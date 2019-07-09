# VBA Form Moderniser
Sets up modern buttons based on userform labels and styles other elements in Office VBA userforms.
This  gives vba userforms a look close to that of modern office forms.

## Before and after

Makes it easy to turn a VBA userform from this:

![Form before](https://github.com/neilt1700/vba-form-moderniser/blob/master/images/sample-form-before.png)

To this:

![Form before](https://github.com/neilt1700/vba-form-moderniser/blob/master/images/sample-form-after.png)

## Summary of what it does
* Creates modern style buttons by converting command buttons into layers of labels.
* Applies general styles to the VBA Userform (as far as that is possible).
* Allows you to use the keyboard to move into the new controls (you need at least one element on your form which can handle keydown events for this to work).

## How to use
You can just download and take a look at [vba-form-moderniser.pptm](https://github.com/neilt1700/vba-form-moderniser/releases). There is a sample form in there and that's called from the MainModule. There is a small amount of extra code you will need to add to your forms as shown in the sample. [Full details are in the Wiki](https://github.com/neilt1700/vba-form-moderniser/wiki/How-to-use-the-VBA-Form-Moderniser).

This now converts command buttons directly. 

## How it works in more detail
The code picks up all command buttons on a form and uses them to create the label controls. Each control is made up of a border layer, background layer, text layer (having the caption from the original command button), and on top, a transparent control layer which receives mouse events and calls the Click method for the original command button (which is now hidden). Mouse up/down/move events are used to control the look of the buttons.

## Office Programmes
While the code is in a PowerPoint file it should work in other Office programmes including Excel and Word. The colour scheme for form controls that appear in the workspace of all these programmes is the same (blue).
