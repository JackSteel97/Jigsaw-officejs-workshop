# Inserting Text
For most stuff in the document Word has the concept of "Ranges". A Range is as it sounds some part of the document with a defined start and end point. Most operations you can execute are done on or relative to a range. The OfficeJs API provides some shortcuts around that, for example you can insert things directly on the document body rather than having to get the range of the body yourself first. 

There are several different ways to insert text into the document, like most things it depends exactly what you want to achieve.

To insert a new paragraph with some text you can use the `insertParagraph` function, notice you have to specify if you want to insert at the `start` or `end` of the range. There is no such thing as the `middle` of the range as an insert location because that would just be the start or end of a different range.

You can also use `insertText`, this one can be used to add text to an existing paragraph like this:
```js
const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
paragraph.insertText(" Hello Jigsaw", Word.InsertLocation.end);
```
Notice that `insertText` will also accept `replace` as an insert location - does what you'd expect and replaces the entire paragraph contents with the new text.

You can also insert textboxes, HTML and raw OOXML, but for now we don't need those.

# Inserting Images
For basic image insertion OfficeJs provides some functions that accept a base64 image. It's important to note that these functions want just the base64 encoded data itself not the label bit (`data:image/jpeg;base64,`) so we have to remove that if present.

Try using `insertInlinePictureFromBase64()` on the paragraph we created earlier.

```js
const base64Img = await downloadImageBase64("https://picsum.photos/200");
const base64DataOnly = base64Img.split(",")[1];
paragraph.insertInlinePictureFromBase64(base64DataOnly, Word.InsertLocation.end);
```

You can also insert floating images and define their size and location. This location is relative to the image's anchor, in this case the inserted paragraph.
```js
const base64Img = await downloadImageBase64("https://picsum.photos/200");
const base64DataOnly = base64Img.split(",")[1];
paragraph.insertPictureFromBase64(base64DataOnly, {top: 100, left: 50});
```

# Searching
Although you can access paragraphs and their text directly and individually if you are working with text and just need to look for specific text for example to find and replace you can use `search` function on the `document.body` object:
```js
const results: Word.RangeCollection = context.document.body.search("Hello World!");
```
This will give you a collection of ranges that match your search term you can iterate to perform any actions against.

You can also search using wildcard patterns:
```js
 const results: Word.RangeCollection = context.document.body.search("$*.[0-9][0-9]", { matchWildcards: true });
```