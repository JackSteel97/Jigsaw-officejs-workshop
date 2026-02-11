# OfficeJs Intro
OfficeJs is the library used for developing modern Office Add-ins. It is a JavaScript library usually loaded from a Microsoft CDN. It enables interaction with the Office document from within our embedded browser runtime.

All Modern Web Add-ins are just webpages that Office opens either in a taskpane view or in a headless browser instance, we can use any of the frameworks and tools usually used for building web applications to create our web add-in.

## Structure of an Office Add-in
The main entry point to your add-in from the perspective of Office is the Manifest. The Manifest defines how things outside of your web page should look (e.g. buttons on the Office Ribbon) and what they should do when clicked. There are only two options there, they can open a taskpane to a specified URL, or they can run a command hosted at a specified URL. We'll cover both of those later.

Most add-ins will use a Taskpane view to display UI, this taskpane URL is set inside the Manifest file, the web page hosted at this location must load and use the OfficeJs library to interact with the Office document in any way.

You can also define a commands URL (it can be the same as the taskpane URL) this is a location the Office Application will open in the background to execute commands that do not need a visible UI to run.

# Setting Up
This codebase was scaffolded using the following steps.

1. Run `npm install -g yo generator-office`
2. Run `yo office`
3. Choose `Office Add-in Task Pane project`
4. Choose `Typescript`
5. Give your add-in a name
6. Select `Word` to support

## Launching and debugging
1. Open the newly created project directory in your terminal or editor and run `npm start`
2. Word will open and you should see your taskpane open or the button available in the Ribbon on the home tab
3. You can debug everything inside the taskpane like a normal web app
	1. To get to the dev tools you can right click -> Inspect
	2. or use the add-in menu popout inside the task pane and click "Attach Debugger"

If you make changes to the manifest file you can check them with `npm validate`

## Tour Scaffolded App
Notice you have a `commands` folder and a `taskpane` folder created inside `src`.
For now you can ignore the `commands` folder, we'll revisit it in the next session.

`taskpane.html` contains some basic HTML for the demo page and includes the link to the office.js CDN script.

Inside `taskpane.ts` is where the magic happens. You can see the `run` function that calls `Word.run` and inserts a paragraph at the end of the body.

This is how almost all interactions with the document happen, inside a `Word.run` and using `context.sync()` to execute.

Sometimes you'll have to call `context.sync()` several times during a single run. Most often you'll discover this need from getting an error when trying to access a property.

For example trying to run something like this:
```js
const paras = context.document.body.paragraphs;
paras.items.forEach((p) => {
  console.log(p);
});
```
You'll get an error.
Instead you need to load items and sync the context before continuing to access the items of the collection:
```js
const paras = context.document.body.paragraphs;
paras.load("items");
await context.sync();
paras.items.forEach((p) => {
  console.log(p);
});
```
