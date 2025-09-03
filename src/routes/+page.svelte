<script lang="ts">
    import { greet } from "$lib/command/greet";
    // import App from "../components/App.svelte";

    let name = $state("");
    let greetMsg = $state("");
</script>

<main>
    <h1>Welcome to Tauri + Svelte</h1>
    <form class="row" onsubmit={async (event) => {event.preventDefault(); greetMsg = await greet(name);}}>
        <input placeholder="Enter a name..." bind:value={name} />
        <button type="submit">Greet</button>
    </form>
    <p>{greetMsg}</p>
    <!-- <App /> -->
</main>
