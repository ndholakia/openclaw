import { normalizeOptionalLowercaseString } from "openclaw/plugin-sdk/text-runtime";
import {
  createChannelReplyPipeline,
  logTypingFailure,
  resolveChannelMediaMaxBytes,
  type OpenClawConfig,
  type MSTeamsReplyStyle,
  type RuntimeEnv,
} from "../runtime-api.js";
import type { MSTeamsAccessTokenProvider } from "./attachments/types.js";
import type { StoredConversationReference } from "./conversation-store.js";
import { emitMSTeamsMessageSentHooks } from "./delivery-hooks.js";
import {
  classifyMSTeamsSendError,
  formatMSTeamsSendErrorHint,
  formatUnknownError,
} from "./errors.js";
import {
  buildConversationReference,
  type MSTeamsAdapter,
  type MSTeamsRenderedMessage,
  renderReplyPayloadsToMessages,
  sendMSTeamsMessages,
} from "./messenger.js";
import type { MSTeamsMonitorLogger } from "./monitor-types.js";
import { createTeamsReplyStreamController } from "./reply-stream-controller.js";
import { withRevokedProxyFallback } from "./revoked-context.js";
import { getMSTeamsRuntime } from "./runtime.js";
import type { MSTeamsTurnContext } from "./sdk-types.js";

export { pickInformativeStatusText } from "./reply-stream-controller.js";

export function createMSTeamsReplyDispatcher(params: {
  cfg: OpenClawConfig;
  agentId: string;
  sessionKey: string;
  accountId?: string;
  runtime: RuntimeEnv;
  log: MSTeamsMonitorLogger;
  adapter: MSTeamsAdapter;
  appId: string;
  conversationRef: StoredConversationReference;
  context: MSTeamsTurnContext;
  replyStyle: MSTeamsReplyStyle;
  textLimit: number;
  onSentMessageIds?: (ids: string[]) => void;
  tokenProvider?: MSTeamsAccessTokenProvider;
  sharePointSiteId?: string;
}) {
  const core = getMSTeamsRuntime();
  const msteamsCfg = params.cfg.channels?.msteams;
  const conversationType = normalizeOptionalLowercaseString(
    params.conversationRef.conversation?.conversationType,
  );
  const isTypingSupported = conversationType === "personal" || conversationType === "groupchat";

  const sendTypingIndicator = isTypingSupported
    ? async () => {
        await withRevokedProxyFallback({
          run: async () => {
            await params.context.sendActivity({ type: "typing" });
          },
          onRevoked: async () => {
            const baseRef = buildConversationReference(params.conversationRef);
            await params.adapter.continueConversation(
              params.appId,
              { ...baseRef, activityId: undefined },
              async (ctx) => {
                await ctx.sendActivity({ type: "typing" });
              },
            );
          },
          onRevokedLog: () => {
            params.log.debug?.("turn context revoked, sending typing via proactive messaging");
          },
        });
      }
    : async () => {};

  const { onModelSelected, typingCallbacks, ...replyPipeline } = createChannelReplyPipeline({
    cfg: params.cfg,
    agentId: params.agentId,
    channel: "msteams",
    accountId: params.accountId,
    typing: {
      start: sendTypingIndicator,
      onStartError: (err) => {
        logTypingFailure({
          log: (message) => params.log.debug?.(message),
          channel: "msteams",
          action: "start",
          error: err,
        });
      },
    },
  });

  const chunkMode = core.channel.text.resolveChunkMode(params.cfg, "msteams");
  const tableMode = core.channel.text.resolveMarkdownTableMode({
    cfg: params.cfg,
    channel: "msteams",
  });
  const mediaMaxBytes = resolveChannelMediaMaxBytes({
    cfg: params.cfg,
    resolveChannelLimitMb: ({ cfg }) => cfg.channels?.msteams?.mediaMaxMb,
  });
  const feedbackLoopEnabled = params.cfg.channels?.msteams?.feedbackEnabled !== false;
  const streamController = createTeamsReplyStreamController({
    conversationType,
    context: params.context,
    feedbackLoopEnabled,
    log: params.log,
  });

  const blockStreamingEnabled =
    typeof msteamsCfg?.blockStreaming === "boolean" ? msteamsCfg.blockStreaming : false;
  const typingIndicatorEnabled =
    typeof msteamsCfg?.typingIndicator === "boolean" ? msteamsCfg.typingIndicator : true;

  const pendingMessages: MSTeamsRenderedMessage[] = [];

  const sendMessages = async (messages: MSTeamsRenderedMessage[]): Promise<string[]> => {
    return sendMSTeamsMessages({
      replyStyle: params.replyStyle,
      adapter: params.adapter,
      appId: params.appId,
      conversationRef: params.conversationRef,
      context: params.context,
      messages,
      retry: {},
      onRetry: (event) => {
        params.log.debug?.("retrying send", {
          replyStyle: params.replyStyle,
          ...event,
        });
      },
      tokenProvider: params.tokenProvider,
      sharePointSiteId: params.sharePointSiteId,
      mediaMaxBytes,
      feedbackLoopEnabled,
    });
  };

  const queueDeliveryFailureSystemEvent = (failure: {
    failed: number;
    total: number;
    error: unknown;
  }) => {
    const classification = classifyMSTeamsSendError(failure.error);
    const errorText = formatUnknownError(failure.error);
    const failedAll = failure.failed >= failure.total;
    const summary = failedAll
      ? "the previous reply was not delivered"
      : `${failure.failed} of ${failure.total} message blocks were not delivered`;
    const sentences = [
      `Microsoft Teams delivery failed: ${summary}.`,
      `The user may not have received ${failedAll ? "that reply" : "the full reply"}.`,
      `Error: ${errorText}.`,
      classification.statusCode != null ? `Status: ${classification.statusCode}.` : undefined,
      classification.kind === "transient" || classification.kind === "throttled"
        ? "Retrying later may succeed."
        : undefined,
    ].filter(Boolean);
    core.system.enqueueSystemEvent(sentences.join(" "), {
      sessionKey: params.sessionKey,
      contextKey: `msteams:delivery-failure:${params.conversationRef.conversation?.id ?? "unknown"}`,
    });
  };

  const flushPendingMessages = async () => {
    if (pendingMessages.length === 0) {
      return;
    }
    const toSend = pendingMessages.splice(0);
    const total = toSend.length;
    let ids: string[];
    let failureCount = 0;
    let lastError: unknown;
    try {
      ids = await sendMessages(toSend);
    } catch (batchError) {
      ids = [];
      lastError = batchError;
      for (const msg of toSend) {
        try {
          const msgIds = await sendMessages([msg]);
          ids.push(...msgIds);
        } catch (msgError) {
          failureCount += 1;
          lastError = msgError;
          params.log.debug?.("individual message send failed, continuing with remaining blocks");
        }
      }
      if (failureCount > 0) {
        params.log.warn?.(`failed to deliver ${failureCount} of ${total} message blocks`, {
          failed: failureCount,
          total,
        });
        queueDeliveryFailureSystemEvent({
          failed: failureCount,
          total,
          error: lastError,
        });
      }
    }
    if (ids.length > 0) {
      params.onSentMessageIds?.(ids);
    }
    // Emit `message:sent` to the internal + plugin-SDK hook buses so
    // downstream listeners (per-user memory loggers, audit substrates) can
    // observe the agent's reply. Mirrors the telegram pattern in
    // `extensions/telegram/src/bot/delivery.replies.ts:emitTelegramMessageSentHooks`.
    const isGroup =
      conversationType === "groupchat" || conversationType === "channel";
    // For personal DMs, the recipient AAD is the canonical "to" — matches
    // telegram's `chatId` semantics. For groups, fall back to conversation id.
    const recipientAad =
      params.conversationRef.user?.aadObjectId ??
      params.conversationRef.aadObjectId;
    const conversationId = params.conversationRef.conversation?.id;
    const to =
      !isGroup && recipientAad ? recipientAad : (conversationId ?? "unknown");
    const content = toSend
      .map((m) => (typeof m.text === "string" ? m.text : ""))
      .filter(Boolean)
      .join("\n\n");
    emitMSTeamsMessageSentHooks({
      sessionKeyForInternalHooks: params.sessionKey,
      to,
      conversationId,
      accountId: params.accountId,
      content,
      success: ids.length > 0,
      error: failureCount > 0 ? formatUnknownError(lastError) : undefined,
      messageId: ids[0],
      isGroup,
      groupId: isGroup ? conversationId : undefined,
    });
  };

  const {
    dispatcher,
    replyOptions,
    markDispatchIdle: baseMarkDispatchIdle,
  } = core.channel.reply.createReplyDispatcherWithTyping({
    ...replyPipeline,
    humanDelay: core.channel.reply.resolveHumanDelayConfig(params.cfg, params.agentId),
    onReplyStart: async () => {
      await streamController.onReplyStart();
      // Avoid duplicate typing UX in DMs: stream status already shows progress.
      if (typingIndicatorEnabled && !streamController.hasStream()) {
        await typingCallbacks?.onReplyStart?.();
      }
    },
    typingCallbacks,
    deliver: async (payload) => {
      const preparedPayload = streamController.preparePayload(payload);
      if (!preparedPayload) {
        return;
      }

      const messages = renderReplyPayloadsToMessages([preparedPayload], {
        textChunkLimit: params.textLimit,
        chunkText: true,
        mediaMode: "split",
        tableMode,
        chunkMode,
      });
      pendingMessages.push(...messages);

      // When block streaming is enabled, flush immediately so blocks are
      // delivered progressively instead of batching until markDispatchIdle.
      if (blockStreamingEnabled) {
        await flushPendingMessages();
      }
    },
    onError: (err, info) => {
      const errMsg = formatUnknownError(err);
      const classification = classifyMSTeamsSendError(err);
      const hint = formatMSTeamsSendErrorHint(classification);
      params.runtime.error?.(
        `msteams ${info.kind} reply failed: ${errMsg}${hint ? ` (${hint})` : ""}`,
      );
      params.log.error("reply failed", {
        kind: info.kind,
        error: errMsg,
        classification,
        hint,
      });
    },
  });

  // Track whether the streaming path has already emitted message:sent for
  // this turn so we don't double-fire when both stream finalize AND the
  // flushPendingMessages fallback ran (e.g. partial-text + media split).
  let streamMessageSentEmitted = false;

  const emitStreamMessageSentIfNeeded = () => {
    if (streamMessageSentEmitted) {
      return;
    }
    if (!streamController.hasStream() || !streamController.isFinalized()) {
      return;
    }
    const content = streamController.streamedContent();
    if (!content) {
      return;
    }
    streamMessageSentEmitted = true;
    const isGroup =
      conversationType === "groupchat" || conversationType === "channel";
    const recipientAad =
      params.conversationRef.user?.aadObjectId ??
      params.conversationRef.aadObjectId;
    const conversationId = params.conversationRef.conversation?.id;
    const to =
      !isGroup && recipientAad ? recipientAad : (conversationId ?? "unknown");
    emitMSTeamsMessageSentHooks({
      sessionKeyForInternalHooks: params.sessionKey,
      to,
      conversationId,
      accountId: params.accountId,
      content,
      success: true,
      isGroup,
      groupId: isGroup ? conversationId : undefined,
    });
  };

  const markDispatchIdle = (): Promise<void> => {
    return flushPendingMessages()
      .catch((err) => {
        const errMsg = formatUnknownError(err);
        const classification = classifyMSTeamsSendError(err);
        const hint = formatMSTeamsSendErrorHint(classification);
        params.runtime.error?.(`msteams flush reply failed: ${errMsg}${hint ? ` (${hint})` : ""}`);
        params.log.error("flush reply failed", {
          error: errMsg,
          classification,
          hint,
        });
      })
      .then(() => {
        return streamController.finalize().catch((err) => {
          params.log.debug?.("stream finalize failed", { error: formatUnknownError(err) });
        });
      })
      .then(() => {
        // After the stream finalizes, emit message:sent for the streamed
        // content so downstream hook handlers (per-user memory loggers,
        // audit substrates) see the agent's reply. Streaming bypasses
        // flushPendingMessages entirely; without this emit, streamed
        // personal-DM replies are silent on the hook bus.
        emitStreamMessageSentIfNeeded();
      })
      .finally(() => {
        baseMarkDispatchIdle();
      });
  };

  return {
    dispatcher,
    replyOptions: {
      ...replyOptions,
      ...(streamController.hasStream()
        ? {
            onPartialReply: (payload: { text?: string }) =>
              streamController.onPartialReply(payload),
          }
        : {}),
      disableBlockStreaming:
        typeof msteamsCfg?.blockStreaming === "boolean" ? !msteamsCfg.blockStreaming : undefined,
      onModelSelected,
    },
    markDispatchIdle,
  };
}
