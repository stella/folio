<script setup lang="ts">
const execCommand = (v, command) => command(v.state, (tr) => v.dispatch(tr), v);
</script>
<template><button>Bold</button></template>
