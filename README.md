# Custom Underline Add-In for Word
This is a C# code for a Microsoft Word Add-In that adds a custom context menu option to underline selected text in a wavy style with a user-specified color.

The code defines an event handler for the DocumentOpen event of the Word application, which adds another event handler for the ContentControlOnEnter event of the document. This event handler adds a custom context menu button to the "Text" command bar of Word.

When the custom button is clicked, the code prompts the user to input the text to be underlined and to select a color for the underline. Then, it searches for all instances of the specified text and applies the wavy underline style with the selected color to each instance.
