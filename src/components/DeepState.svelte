<script>
    const numbers = $state([1, 2, 3, 4]);
    const total = $derived(numbers.reduce((acc, val) => acc + val, 0));

    function addNumber() {
        numbers.push(numbers.length + 1);
        // console.log($state.snapshot(numbers));
    }
    $inspect(numbers).with(console.log);
</script>

<p>{numbers.join(" + ")} = {total}</p>

<button class="btn btn-soft" onclick={addNumber}> Add a number </button>
