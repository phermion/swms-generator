<template>
  <md-card>
    <!-- Heading -->
    <md-card-header>
      <div class="md-title">SWMS Form Generator</div>
      <div class="md-subhead">Generate them forms</div>
    </md-card-header>

    <!-- Form -->
    <md-card-content>
      <form>
        <div class="md-layout md-gutter">
          <!-- First Column -->
          <div class="md-layout-item">
            <md-field>
              <label>Select Template File (.docx)</label>
              <md-file @change="onFileSelect($event.target.name, $event.target.files)" accept=".docx" />
            </md-field>
          </div>
          <!-- Second Column -->
          <div class="md-layout-item">
            <md-field>
              <label>Project Name</label>
              <md-input v-model="templateData.project.name"></md-input>
            </md-field>
            <md-field class="has-danger">
              <label>Contractor</label>
              <md-input v-model="templateData.project.contractor"></md-input>
            </md-field>
          </div>
          <!-- Third Column -->
          <div class="md-layout-item">
            <!-- List -->
            <md-field>
              <label for="movies">Bullet Point List</label>
              <md-select v-model="templateData.bulletPointList" name="bullet-point-list" id="bullet-point-list" multiple>
                <md-option value="First option text that will appear in document">First Option</md-option>
                <md-option value="Second option text that will appear in document">Second Option</md-option>
                <md-option value="Third option text that will appear in document">Third Option</md-option>
              </md-select>
            </md-field>
            <!-- Checkbox -->
            <md-checkbox v-model="templateData.checkbox" class="md-primary">Checkbox</md-checkbox>
          </div>
        </div>
        <md-button @click="onSubmit()" class="md-raised md-primary">
          Generate
        </md-button>
      </form>
    </md-card-content>
  </md-card>
</template>

<script>
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

export default {
  name: 'Form',
  props: {
    msg: String
  },
  data () {
    return {
      templateFileName: '',
      templateFile: [],
      templateData: {
        project: {
          name: '',
          contractor: '',
        },
        bulletPointList: [],
        checkbox: false
      }
    }
  },
  methods: {
    onFileSelect(fileName, files) {
      this.templateFileName = fileName;
      this.templateFile = files[0];
      console.log(this.templateFile)
    },
    onSubmit() {
      this.generateWordDoc(this.templateData);
    },
    generateWordDoc(templateData) {
      // https://docxtemplater.readthedocs.io/en/latest/generate.html#browser
      if (!this.templateFile) {
        return;
      }

      var reader = new FileReader();
      reader.readAsBinaryString(this.templateFile);
      reader.onload = (event) => {
          const content = event.target.result;
          var zip = new PizZip(content);
          let doc = new window.docxtemplater(zip);

          try {
            doc = new window.docxtemplater(zip);
          }
          catch (error) {
              this.errorHandler(error);
          }

          doc.setData(templateData);

          try {
              doc.render();
          }
          catch (error) {
              this.errorHandler(error);
          }

          var outputDocument = doc.getZip().generate({
              type:"blob",
              mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          })

          saveAs(outputDocument,"output.docx")
      }
    },
    errorHandler(error) {
      // https://docxtemplater.readthedocs.io/en/latest/generate.html#browser
      console.log(JSON.stringify({error: error}));

      if (error.properties && error.properties.errors instanceof Array) {
          const errorMessages = error.properties.errors.map(function (error) {
              return error.properties.explanation;
          }).join("\n");
          console.log('errorMessages', errorMessages);
      }
    },
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
