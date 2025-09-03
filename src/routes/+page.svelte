<script lang="ts">
    import { invoke } from "@tauri-apps/api/core";
    import App from "../components/App.svelte";

    let name = $state("");
    let greetMsg = $state("");

    async function greet(event: Event) {
        event.preventDefault();
        // Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
        greetMsg = await invoke("greet", { name });
    }
</script>

<main>
    <h1>Welcome to Tauri + Svelte</h1>
    <form class="row" onsubmit={greet}>
        <input placeholder="Enter a name..." bind:value={name} />
        <button type="submit">Greet</button>
    </form>
    <p>{greetMsg}</p>
    <!-- <App /> -->
</main>
