import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import {saveAs} from "file-saver";
let PizZipUtils : any = null;
if (typeof window !== "undefined") {
  import("pizzip/utils/index.js").then(function (r) {
    PizZipUtils = r;
  });
}

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

const generateDocument = () => {
  loadFile("/template.docx", (error, content) => {
    if(error){
      throw error
    }

    let zip = new PizZip(content);
    let doc = new Docxtemplater().loadZip(zip);

    doc.setData({
      expno : "01",
      expname : `Verification of Kirchhoff’s Voltage Law
      and Voltage Divider Rule in a Series Circuit`,
      oddeven : "Even",
      groupnum : "B-05",
      expDate : "13/11/2022",
      subDate : "01/12/2022",
      student1 : "Mominul Islam Hemal(AE-40)",
      student2: "Talha Jubair Siam(FH-52)",
      student3 : "Rezaunnabi Ruhan(AE-58)"
    });

    try {
      doc.render();
    } catch (error : any) {
      const replaceErrors = (key, value) => {
        if (value instanceof Error) {
          return Object.getOwnPropertyNames(value).reduce(function (
            error,
            key
          ) {
            error[key] = value[key];
            return error;
          },
          {});
        }
        return value;
      }

      console.log(JSON.stringify({ error: error }, replaceErrors));

      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors
          .map(function (error) {
            return error.properties.explanation;
          })
          .join("\n");
        console.log("errorMessages", errorMessages);
        // errorMessages is a humanly readable message looking like this :
        // 'The tag beginning with "foobar" is unopened'
      }

      throw error;
    }

    console.log("aaa");
    var out = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }); 
    // Output the document using Data-URI
    saveAs(out, "labCover.docx");
    console.log("aaa2");
  })
}

const Homepage = () => {
  return (
    <div>
      <button onClick={generateDocument}>Generate</button>
    </div>
  )
}


export default Homepage;