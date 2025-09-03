<script>
    let count = $state(0);

    function increment() {
        count += 1;
    }
</script>

<div class="m-2 p-2">
    <button class="btn btn-soft btn-success" onclick={increment}>
        Clicked {count}
        {count === 1 ? "time" : "times"}
    </button>
</div>
