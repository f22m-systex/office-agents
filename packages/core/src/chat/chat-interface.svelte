<script lang="ts">
  import { type AgentContext, getSessionMessageCount } from "@office-agents/sdk";
  import {
    Check,
    ChevronDown,
    Contrast,
    Copy,
    Download,
    Eye,
    EyeOff,
    FolderOpen,
    MessageSquare,
    Moon,
    Plus,
    Settings,
    Sun,
    Trash2,
    Upload,
  } from "lucide-svelte";
  import { onDestroy } from "svelte";
  import type { AppAdapter } from "./app-adapter";
  import { ChatController } from "./chat-controller";
  import { setChatContext } from "./chat-runtime-context";
  import ChatInput from "./chat-input.svelte";
  import FilesPanel from "./files-panel.svelte";
  import MessageList from "./message-list.svelte";
  import SettingsPanel from "./settings-panel.svelte";
  import type { ChatTab } from "./types";
  import {
    chatMessagesToJson,
    chatMessagesToMarkdown,
    copyToClipboard,
    downloadFile,
    sanitizeFilename,
  } from "./export-utils";

  type Theme = "light" | "dark";

  const THEME_KEY = "office-agents-theme";
  const HIGH_CONTRAST_KEY = "office-agents-high-contrast";

  interface Props {
    adapter: AppAdapter;
    context: AgentContext;
  }

  let { adapter, context }: Props = $props();

  const controller = (() => {
    const ctrl = new ChatController(adapter, context);
    setChatContext(ctrl);
    return ctrl;
  })();

  const runtimeState = controller.state;

  let activeTab = $state<ChatTab>("chat");
  let isDragOver = $state(false);
  let dragCounter = $state(0);
  let sessionDropdownOpen = $state(false);
  let sessionDropdownRef = $state<HTMLDivElement | null>(null);
  let exportDropdownOpen = $state(false);
  let exportDropdownRef = $state<HTMLDivElement | null>(null);
  let copied = $state(false);

  let theme = $state<Theme>(loadTheme());
  let highContrast = $state<boolean>(loadHighContrast());

  function loadTheme(): Theme {
    const saved = localStorage.getItem(THEME_KEY) as Theme | null;
    const initial =
      saved ??
      (window.matchMedia("(prefers-color-scheme: light)").matches
        ? "light"
        : "dark");
    document.documentElement.setAttribute("data-theme", initial);
    return initial;
  }

  function loadHighContrast(): boolean {
    const saved = localStorage.getItem(HIGH_CONTRAST_KEY);
    return saved === "true";
  }

  function toggleTheme() {
    theme = theme === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", theme);
    localStorage.setItem(THEME_KEY, theme);
  }

  function toggleHighContrast() {
    highContrast = !highContrast;
    if (highContrast) {
      document.documentElement.setAttribute("data-theme-mode", "high-contrast");
    } else {
      document.documentElement.removeAttribute("data-theme-mode");
    }
    localStorage.setItem(HIGH_CONTRAST_KEY, highContrast ? "true" : "false");
  }

  function formatTokens(value: number): string {
    if (value >= 1_000_000) return `${(value / 1_000_000).toFixed(1)}M`;
    if (value >= 1_000) return `${(value / 1_000).toFixed(1)}k`;
    return value.toString();
  }

  function formatCost(value: number): string {
    if (value < 0.01) return `$${value.toFixed(4)}`;
    return `$${value.toFixed(3)}`;
  }

  async function handleNewSession() {
    await controller.newSession();
    sessionDropdownOpen = false;
    activeTab = "chat";
  }

  async function handleSwitchSession(sessionId: string) {
    await controller.switchSession(sessionId);
    sessionDropdownOpen = false;
    activeTab = "chat";
  }

  function handleClickOutside(event: MouseEvent) {
    const target = event.target;
    if (!(target instanceof Node)) return;

    if (sessionDropdownRef && !sessionDropdownRef.contains(target)) {
      sessionDropdownOpen = false;
    }
    if (exportDropdownRef && !exportDropdownRef.contains(target)) {
      exportDropdownOpen = false;
    }
  }

  async function handleCopyMarkdown() {
    const messages = $runtimeState.messages;
    const md = chatMessagesToMarkdown(messages, currentName);
    const success = await copyToClipboard(md);
    if (success) {
      copied = true;
      setTimeout(() => {
        copied = false;
      }, 2000);
    }
  }

  function handleExportMarkdown() {
    const messages = $runtimeState.messages;
    const md = chatMessagesToMarkdown(messages, currentName);
    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `chat_${sanitizeFilename(currentName)}_${dateStr}.md`;
    downloadFile(md, filename, "text/markdown");
    exportDropdownOpen = false;
  }

  function handleExportJson() {
    const messages = $runtimeState.messages;
    const json = chatMessagesToJson(messages, currentName);
    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `chat_${sanitizeFilename(currentName)}_${dateStr}.json`;
    downloadFile(json, filename, "application/json");
    exportDropdownOpen = false;
  }

  function handleDragEnter(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    dragCounter += 1;
    if (event.dataTransfer?.types.includes("Files")) {
      isDragOver = true;
    }
  }

  function handleDragLeave(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    dragCounter -= 1;
    if (dragCounter === 0) {
      isDragOver = false;
    }
  }

  function handleDragOver(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
  }

  function handleDrop(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    dragCounter = 0;
    isDragOver = false;

    const files = Array.from(event.dataTransfer?.files ?? []);
    if (files.length > 0) {
      void controller.processFiles(files);
    }
  }

  $effect(() => {
    controller.setAdapter(adapter);
  });

  $effect(() => {
    if (!sessionDropdownOpen && !exportDropdownOpen) return undefined;
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  });

  $effect(() => {
    const root = document.documentElement;
    const localStoragePrefix = context.namespace.localStoragePrefix;
    const family = localStorage.getItem(`${localStoragePrefix}-font-family`) || "default";
    const size = localStorage.getItem(`${localStoragePrefix}-font-size`) || "14px";

    root.style.fontSize = size;
    if (family !== "default") {
      root.style.setProperty("--chat-font-mono", family);
    } else {
      root.style.removeProperty("--chat-font-mono");
    }
  });

  onDestroy(() => {
    controller.dispose();
  });

  const currentName = $derived(
    $runtimeState.currentSession?.name ?? "New Chat",
  );
  const truncatedName = $derived(
    currentName.length > 20 ? `${currentName.slice(0, 18)}…` : currentName,
  );
  const followMode = $derived($runtimeState.providerConfig?.followMode ?? true);
  const HeaderExtras = $derived(adapter.HeaderExtras);
  const SelectionIndicator = $derived(adapter.SelectionIndicator);
</script>

<div
  role="application"
  class="flex flex-col h-full bg-(--chat-bg) relative"
  style="font-family: var(--chat-font-mono)"
  ondragenter={handleDragEnter}
  ondragleave={handleDragLeave}
  ondragover={handleDragOver}
  ondrop={handleDrop}
>
  <div class="border-b border-(--chat-border) bg-(--chat-bg)">
    <div class="flex items-center justify-between px-2">
      <div class="flex">
        {#if activeTab === "chat"}
          <div class="relative" bind:this={sessionDropdownRef}>
            <button
              type="button"
              onclick={() => (sessionDropdownOpen = !sessionDropdownOpen)}
              class="flex items-center gap-1 px-3 py-2 text-xs uppercase tracking-wider border-b-2 border-(--chat-accent) text-(--chat-text-primary) transition-colors"
              style="font-family: var(--chat-font-mono)"
            >
              <MessageSquare size={12} />
              <span class="max-w-[100px] truncate">{truncatedName}</span>
              <ChevronDown
                size={12}
                class={`transition-transform ${sessionDropdownOpen ? "rotate-180" : ""}`}
              />
            </button>

            {#if sessionDropdownOpen}
              <div
                class="absolute top-full left-0 mt-1 w-56 bg-(--chat-bg) border border-(--chat-border) rounded shadow-lg z-50 overflow-hidden"
              >
                <button
                  type="button"
                  onclick={handleNewSession}
                  disabled={$runtimeState.isStreaming}
                  class={`w-full flex items-center gap-2 px-3 py-2 text-xs transition-colors border-b border-(--chat-border) ${$runtimeState.isStreaming ? "text-(--chat-text-muted) cursor-not-allowed" : "text-(--chat-accent) hover:bg-(--chat-bg-secondary)"}`}
                >
                  <Plus size={14} />
                  New Chat
                </button>

                <div class="max-h-48 overflow-y-auto">
                  {#each $runtimeState.sessions as session (session.id)}
                    {@const isCurrent =
                      session.id === $runtimeState.currentSession?.id}
                    {@const isDisabled =
                      $runtimeState.isStreaming && !isCurrent}
                    <button
                      type="button"
                      disabled={isDisabled}
                      class={`flex items-center justify-between px-3 py-2 text-xs transition-colors w-full text-left ${isCurrent ? "bg-(--chat-bg-secondary)" : ""} ${isDisabled ? "cursor-not-allowed opacity-50" : "cursor-pointer hover:bg-(--chat-bg-secondary)"}`}
                      onclick={() => handleSwitchSession(session.id)}
                    >
                      <div class="flex items-center gap-2 min-w-0 flex-1">
                        {#if isCurrent}
                          <Check
                            size={12}
                            class="text-(--chat-accent) shrink-0"
                          />
                        {:else}
                          <div class="w-3 shrink-0"></div>
                        {/if}
                        <span class="truncate text-(--chat-text-primary)">
                          {session.name}
                        </span>
                      </div>
                      <span
                        class="text-[10px] text-(--chat-text-muted) shrink-0 ml-2"
                      >
                        {getSessionMessageCount(session)}
                      </span>
                    </button>
                  {/each}
                </div>

                {#if $runtimeState.sessions.length > 1 && $runtimeState.currentSession}
                  <button
                    type="button"
                    disabled={$runtimeState.isStreaming}
                    onclick={async (event) => {
                      event.stopPropagation();
                      await controller.deleteCurrentSession();
                      sessionDropdownOpen = false;
                    }}
                    class={`w-full flex items-center gap-2 px-3 py-2 text-xs transition-colors border-t border-(--chat-border) ${$runtimeState.isStreaming ? "text-(--chat-text-muted) cursor-not-allowed" : "text-(--chat-error) hover:bg-(--chat-bg-secondary)"}`}
                  >
                    <Trash2 size={14} />
                    Delete Current Session
                  </button>
                {/if}
              </div>
            {/if}
          </div>
        {:else}
          <button
            type="button"
            onclick={() => (activeTab = "chat")}
            class="flex items-center gap-1.5 px-3 py-2 text-xs uppercase tracking-wider border-b-2 border-transparent transition-colors text-(--chat-text-muted) hover:text-(--chat-text-secondary)"
            style="font-family: var(--chat-font-mono)"
          >
            <MessageSquare size={12} />
            Chat
          </button>
        {/if}

        <button
          type="button"
          onclick={() => (activeTab = "files")}
          class={`flex items-center gap-1.5 px-3 py-2 text-xs uppercase tracking-wider border-b-2 transition-colors ${activeTab === "files" ? "border-(--chat-accent) text-(--chat-text-primary)" : "border-transparent text-(--chat-text-muted) hover:text-(--chat-text-secondary)"}`}
          style="font-family: var(--chat-font-mono)"
        >
          <FolderOpen size={12} />
          Files
        </button>

        <button
          type="button"
          onclick={() => (activeTab = "settings")}
          class={`flex items-center gap-1.5 px-3 py-2 text-xs uppercase tracking-wider border-b-2 transition-colors ${activeTab === "settings" ? "border-(--chat-accent) text-(--chat-text-primary)" : "border-transparent text-(--chat-text-muted) hover:text-(--chat-text-secondary)"}`}
          style="font-family: var(--chat-font-mono)"
        >
          <Settings size={12} />
          Settings
        </button>
      </div>

      <div class="flex items-center">
        {#if activeTab === "chat" && HeaderExtras}
          <HeaderExtras />
        {/if}

        {#if activeTab === "chat" && (adapter.showFollowModeToggle ?? true)}
          <button
            type="button"
            onclick={() => controller.toggleFollowMode()}
            class={`p-1.5 transition-colors ${followMode ? "text-(--chat-accent) hover:text-(--chat-text-primary)" : "text-(--chat-text-muted) hover:text-(--chat-text-primary)"}`}
            data-tooltip={followMode ? "Follow mode: ON" : "Follow mode: OFF"}
          >
            {#if followMode}
              <Eye size={14} />
            {:else}
              <EyeOff size={14} />
            {/if}
          </button>
        {/if}

        <button
          type="button"
          onclick={toggleTheme}
          class="p-1.5 text-(--chat-text-muted) hover:text-(--chat-text-primary) transition-colors"
          data-tooltip={theme === "dark" ? "Light mode" : "Dark mode"}
        >
          {#if theme === "dark"}
            <Sun size={14} />
          {:else}
            <Moon size={14} />
          {/if}
        </button>

        <button
          type="button"
          onclick={toggleHighContrast}
          class={`p-1.5 transition-colors ${highContrast ? "text-(--chat-accent)" : "text-(--chat-text-muted) hover:text-(--chat-text-primary)"}`}
          data-tooltip={highContrast ? "High contrast: ON" : "High contrast: OFF"}
        >
          <Contrast size={14} />
        </button>

        {#if activeTab === "chat" && $runtimeState.messages.length > 0}
          <div class="relative flex items-center" bind:this={exportDropdownRef}>
            <button
              type="button"
              onclick={() => (exportDropdownOpen = !exportDropdownOpen)}
              class={`p-1.5 transition-colors ${exportDropdownOpen ? "text-(--chat-accent)" : "text-(--chat-text-muted) hover:text-(--chat-accent)"}`}
              data-tooltip="Export or copy history"
            >
              <Download size={14} />
            </button>

            {#if exportDropdownOpen}
              <div
                class="absolute top-full right-0 mt-1 w-52 bg-(--chat-bg) border border-(--chat-border) rounded shadow-lg z-50 overflow-hidden"
              >
                <button
                  type="button"
                  onclick={handleCopyMarkdown}
                  class="w-full flex items-center gap-2 px-3 py-2 text-xs text-(--chat-text-primary) hover:bg-(--chat-bg-secondary) text-left transition-colors cursor-pointer"
                >
                  {#if copied}
                    <Check size={12} class="text-(--chat-accent) shrink-0" />
                    <span class="truncate text-(--chat-accent)">Copied!</span>
                  {:else}
                    <Copy size={12} class="shrink-0" />
                    <span class="truncate">Copy as Markdown</span>
                  {/if}
                </button>
                <button
                  type="button"
                  onclick={handleExportMarkdown}
                  class="w-full flex items-center gap-2 px-3 py-2 text-xs text-(--chat-text-primary) hover:bg-(--chat-bg-secondary) text-left transition-colors cursor-pointer border-t border-(--chat-border)"
                >
                  <Download size={12} class="shrink-0" />
                  <span class="truncate">Download Markdown (.md)</span>
                </button>
                <button
                  type="button"
                  onclick={handleExportJson}
                  class="w-full flex items-center gap-2 px-3 py-2 text-xs text-(--chat-text-primary) hover:bg-(--chat-bg-secondary) text-left transition-colors cursor-pointer border-t border-(--chat-border)"
                >
                  <Download size={12} class="shrink-0" />
                  <span class="truncate">Download JSON (.json)</span>
                </button>
              </div>
            {/if}
          </div>

          <button
            type="button"
            onclick={() => controller.clearMessages()}
            class="p-1.5 text-(--chat-text-muted) hover:text-(--chat-error) transition-colors"
            data-tooltip="Clear messages"
          >
            <Trash2 size={14} />
          </button>
        {/if}
      </div>
    </div>
  </div>

  {#if activeTab === "chat"}
    <MessageList />
    {#if SelectionIndicator}
      <SelectionIndicator />
    {/if}
    <ChatInput />
    {#if $runtimeState.providerConfig}
      <div
        class="flex items-center justify-between px-3 py-1.5 text-[10px] border-t border-(--chat-border) bg-(--chat-bg-secondary) text-(--chat-text-muted)"
        style="font-family: var(--chat-font-mono)"
      >
        <div class="flex items-center gap-3">
          <span title="Input tokens">
            ↑{formatTokens($runtimeState.sessionStats.inputTokens)}
          </span>
          <span title="Output tokens">
            ↓{formatTokens($runtimeState.sessionStats.outputTokens)}
          </span>
          {#if $runtimeState.sessionStats.cacheRead > 0}
            <span title="Cache read tokens">
              R{formatTokens($runtimeState.sessionStats.cacheRead)}
            </span>
          {/if}
          {#if $runtimeState.sessionStats.cacheWrite > 0}
            <span title="Cache write tokens">
              W{formatTokens($runtimeState.sessionStats.cacheWrite)}
            </span>
          {/if}
          <span title="Total cost">
            {formatCost($runtimeState.sessionStats.totalCost)}
          </span>
          {#if $runtimeState.sessionStats.contextWindow > 0}
            <span title="Context usage">
              {(
                (($runtimeState.sessionStats.lastInputTokens || 0) /
                  $runtimeState.sessionStats.contextWindow) *
                100
              ).toFixed(1)}%/{formatTokens(
                $runtimeState.sessionStats.contextWindow,
              )}
            </span>
          {/if}
        </div>
        <div class="flex items-center gap-1">
          <span>{$runtimeState.providerConfig.provider}</span>
          <span class="text-(--chat-text-secondary)">
            {$runtimeState.providerConfig.model}
          </span>
          {#if $runtimeState.providerConfig.thinking !== "none"}
            <span class="text-(--chat-accent)">
              • {$runtimeState.providerConfig.thinking}
            </span>
          {/if}
        </div>
      </div>
    {/if}
  {:else if activeTab === "files"}
    <FilesPanel />
  {:else}
    <SettingsPanel />
  {/if}

  {#if isDragOver}
    <div
      class="absolute inset-0 z-50 flex items-center justify-center bg-(--chat-bg)/80 backdrop-blur-sm"
    >
      <div
        class="flex flex-col items-center gap-3 p-8 border-2 border-dashed border-(--chat-accent) rounded-lg"
      >
        <Upload size={32} class="text-(--chat-accent)" />
        <span class="text-sm text-(--chat-text-primary)">Drop files here</span>
      </div>
    </div>
  {/if}
</div>
