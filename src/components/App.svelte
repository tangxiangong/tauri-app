<script lang="ts">
    import Counter from "./Counter.svelte";
    import DeepState from "./DeepState.svelte";
    import Nested from "./Nested.svelte";
    import Props from "./Props.svelte";
    import State from "./State.svelte";

    const name = "Svelte";
    const icon = "src/lib/assets/favicon.svg";
    const string = `This string contains some <strong>HTML!!!</strong>`;
</script>

<p>This is a paragraph.</p>
<Nested />
<p>{@html string}</p>
<State />
<DeepState />

<Counter />
<Counter />
<Counter />

<Props />
<Props answer={12} />

<style>
    p {
        color: goldenrod;
        font-family: "Comic Sans MS", cursive;
        font-size: 2em;
    }
</style>
