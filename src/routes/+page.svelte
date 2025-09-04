<script lang="ts">
    import { greet } from "$lib/command/greet";
    import { selectFile } from "$lib/command/utils";
    // import App from "../components/App.svelte";

    let name = $state("");
    let greetMsg = $state("");
    let selectedFile = $state("");
</script>

<main>
    <h1>Welcome to Tauri + Svelte</h1>
    <form
        class="row"
        onsubmit={async (event) => {
            event.preventDefault();
            greetMsg = await greet(name);
        }}
    >
        <input placeholder="Enter a name..." bind:value={name} />
        <button type="submit">Greet</button>
    </form>
    <p>{greetMsg}</p>

    <button
        class="btn btn-soft btn-success"
        onclick={async () => {
            const file = await selectFile();
            file && (selectedFile = file);
        }}
    >
        Select File
    </button>

    <p>{selectedFile}</p>

    <!-- <App /> -->
</main>
