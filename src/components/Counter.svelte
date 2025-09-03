<script lang="ts">
    import { counter } from "$lib/shared.svelte";
</script>

<div>
    <button class="btn btn-primary" onclick={() => (counter.count += 1)}>
        clicks: {counter.count}
    </button>
</div>
