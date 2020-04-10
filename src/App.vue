<template>
  <div id="app">
    <!-- <img alt="Vue logo" src="./assets/logo.png"> -->
    <Form />
  </div>
</template>

<script>
import Form from './components/Form.vue'

export default {
  name: 'App',
  components: {
    Form
  }
}
</script>

<style lang="scss">
#app {
  max-width: 1240px;
  margin: 100px auto auto auto;
}
</style>
