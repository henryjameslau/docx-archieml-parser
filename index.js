import {convertToHtml} from "mammoth"
import {DomHandler,Parser} from "htmlparser2"
import archieml from "archieml"
import {writeFileSync} from 'fs'

function extractText(elements) {
  let text = "";
  for (const element of elements) {
      if (element.type === "text") {
          text += element.data;
      } else if (element.type === "tag" && element.name !== "a") {
          const childText = extractText(element.children);
          // Add a line break after child text, except for the first element
          if (text.length > 0) {
              text += "\n";
          }
          text += childText;
      } else if (element.type === "tag" && element.name === "a") {
          text += `<a${stringifyAttributes(element.attribs)}>${extractText(element.children)}</a>`;
      }
  }
  return text;
}

function stringifyAttributes(attribs) {
  return Object.entries(attribs)
      .map(([key, value]) => ` ${key}="${value}"`).join("");
}

const handler = new DomHandler((error,dom)=>{
  if (error){
    console.log(error)
  }else{
    const extractedText = extractText(dom);
    console.log(archieml.load(extractedText)); 
    writeFileSync('text.json',JSON.stringify({text:archieml.load(extractedText)}),(err)=>{
      if(err){
        console.error("error saving file:",err)
      } else {
        console.log("text saved to text.json")
      }
    })
  }
})

const parser2 = new Parser(handler)

convertToHtml({ path: "./demo.docx" })
  .then((result) => {
    parser2.write(result.value)
    parser2.end()
  })



