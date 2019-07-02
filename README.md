# vba-form-moderniser
Sets up modern buttons based on userform labels and styles other elements in userforms.
This  gives vba userforms a look close to that of modern office forms.

## Summary of what it does
* Creates modern style buttons from labels.
* Applies general styles to the VBA Userform (as far as that is possible).
* Allows you to use the keyboard to move into the new controls (you need at least one element on your form which can handle keydown events for this to work).

## How to use
You can just download and take a look at vba-form-moderniser.pptm. There is a sample form in there and that's called from the MainModule. There is a small amount of extra code you will need to add to your forms as shown in the sample.

You should also replace any command buttons with labels, the caption on the label will be the text on the new style command button that gets created when you run the code. You don't need to do any formatting on the label or any text positioning - that gets done for you. Your "label buttons" should have a click event for running whatever code you want it to run.

## How it works in more detail
The code picks up all labels which have a specific prefix to their names ("LabelButton") and uses those to create the label controls. Each control is made up of a border layer, background layer, text layer (having the caption from the original label), and on top, a transparent control layer (which receives mouse events and being the original label). Mouse up/down/move events are used to control the look of the buttons.

## Office Programmes
While the code is in a PowerPoint file it should work other Office programmes including Excel and Word. The colour scheme for form controls that appear in the workspace of all these programmes is the same (blue).
