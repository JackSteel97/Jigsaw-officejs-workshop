Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const findInput = document.getElementById("find-input") as HTMLInputElement;
  const replaceInput = document.getElementById("replace-input") as HTMLInputElement;
  console.debug("Find value", findInput.value);
  console.debug("Replace value", replaceInput.value);
  
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // Everything below here is part of the boilerplate code created by scaffolding the app and can be deleted whenever you want.
    
    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
