<script lang="ts">
  import { Code, Download, Eye, Images } from "lucide-svelte";
  import { renderMermaid, svgToBase64 } from "./mermaid-utils";
  import { getChatContext } from "./chat-runtime-context";

  interface Props {
    code: string;
  }

  let { code }: Props = $props();
  const chat = getChatContext();

  let svg = $state("");
  let error = $state("");
  let showCode = $state(false);
  let isInserting = $state(false);

  const id = `mermaid-${Math.random().toString(36).slice(2, 9)}`;

  $effect(() => {
    renderMermaid(code, id)
      .then((res) => {
        svg = res;
        error = "";
      })
      .catch((err) => {
        error = err instanceof Error ? err.message : "Failed to render Mermaid";
        svg = "";
      });
  });

  async function handleInsert() {
    if (!svg || !chat.adapter.insertImage || isInserting) return;

    isInserting = true;
    try {
      const base64 = svgToBase64(svg);
      await chat.adapter.insertImage(base64, "image/svg+xml");
    } catch (err) {
      console.error("Failed to insert Mermaid chart:", err);
    } finally {
      isInserting = false;
    }
  }
</script>

<div class="mermaid-block my-2 border border-(--chat-border) bg-(--chat-bg-secondary)" style="border-radius: var(--chat-radius)">
  <div class="flex items-center justify-between px-2 py-1 border-b border-(--chat-border) bg-(--chat-bg)">
    <span class="text-[10px] uppercase tracking-wider text-(--chat-text-muted) font-mono">Mermaid Diagram</span>
    <div class="flex items-center gap-1">
      <button
        type="button"
        onclick={() => (showCode = !showCode)}
        class="p-1 text-(--chat-text-muted) hover:text-(--chat-text-primary) transition-colors"
        title={showCode ? "View Diagram" : "View Code"}
      >
        {#if showCode}
          <Eye size={12} />
        {:else}
          <Code size={12} />
        {/if}
      </button>

      {#if chat.adapter.insertImage && svg && !showCode}
        <button
          type="button"
          onclick={handleInsert}
          disabled={isInserting}
          class="p-1 text-(--chat-text-muted) hover:text-(--chat-accent) transition-colors disabled:opacity-50"
          title="Insert to Document"
        >
          <Images size={12} class={isInserting ? "animate-pulse" : ""} />
        </button>
      {/if}
    </div>
  </div>

  <div class="p-3 overflow-auto max-h-[400px]">
    {#if error}
      <div class="text-(--chat-error) text-xs font-mono whitespace-pre-wrap">{error}</div>
    {:else if showCode}
      <pre class="text-xs font-mono m-0 text-(--chat-text-primary)"><code>{code}</code></pre>
    {:else if svg}
      <div class="flex justify-center bg-white p-2 rounded">
        {@html svg}
      </div>
    {:else}
      <div class="flex justify-center py-4">
        <div class="animate-pulse text-(--chat-text-muted) text-xs">Rendering...</div>
      </div>
    {/if}
  </div>
</div>

<style>
  .mermaid-block :global(svg) {
    max-width: 100%;
    height: auto;
  }
</style>
