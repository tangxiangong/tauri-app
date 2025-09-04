<script lang="ts">
    import { isSupported } from "$lib/command";
    import Memory from "./Memory.svelte";
</script>

{#await isSupported() then is_supported}
    {#if is_supported}
        <Memory />
    {:else}
        <div role="alert" class="alert alert-error w-1/3">
            <img src="/error.svg" alt="failed" width="20" />
            <span class="font-bold"
                >Unsupported! Only Android, FreeBSD, iOS, Linux, macOS Raspberry
                Pi, and Windows Supported</span
            >
        </div>
    {/if}
{/await}
