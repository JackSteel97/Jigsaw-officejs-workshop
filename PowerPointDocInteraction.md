Powerpoint is quite different to Word in how most things are handled, there are some similarities at a high level but PowerPoint works much more with the concept of "Shapes". Almost everything visible in a PowerPoint slide is some variant of a Shape.

# Differences from Word
A lot of aspects of the API for PowerPoint are actually better thought out and complete than those for Word, its a lot easier to work with and do complex stuff in PowerPoint than the equivalent in Word.

The main difference and the one that you'll run into if you try to share any code between Word and PPT is that you need to use `Word.run` in a Word add-in and `PowerPoint.run` in a PowerPoint document. This means that you effectively have to split any code that interacts with the document out into two separate versions. These bits of code may share internal functions that don't need to do anything with the document directly but you'll need to implement a strategy pattern and detect which host you're running in.

## Sharing Code further reading
See `src/services/office/platformManager.ts` in the Office.VueSPA project to see how we detect the difference between Word/PowerPoint and others in production.

In practice you will likely want a shared interface that your UI code can use to make calls to change things in the document without worrying about if it's Word or PowerPoint, it is a good idea to try and isolate any code that needs to use parts of OfficeJs as much as possible from your business logic, both because it makes switching between hosts easier and also because OfficeJs code can be a pain to Unit Test against.

# Working with Shapes
Most things in PowerPoint are a subset of Shape, this includes Lines, Tables, Text Boxes, and any generic Geometric Shapes. All these shapes are accessible from each slide via the shapes collection.

You can read existing shapes:
```js
const shapeById = slide.shapes.getItem("12");

const allShapes = wholeDocumentRange.shapes;
allShapes.load();
await context.sync();
```

Or add new shapes:
```js
const rectImgContainer = slide.shapes.addGeometricShape(
	PowerPoint.GeometricShapeType.rectangle,
	{
	  top: 0,
	  left: 0,
	  width: 192,
	  height: 72,
	}
);
```

Shapes can also be tagged with key-value pairs that can be used to attach arbitrary string data to the shape for later use:
```js
shape.tags.add("SOME-KEY", "some-value");
```
Be aware: tag keys must always be upper case because of a known bug with PowerPoint (https://github.com/OfficeDev/office-js/issues/6079)

