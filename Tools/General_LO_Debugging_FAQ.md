[This is a comment, written in MarkDown formatting]: #

[This is my personal reference for issues I've had with LibreOffice, and how I've fixed them. Maybe this will help someone else too! :) ]: #
# LibreOffice General/Regular Text

  * ## Efficiency
    * ### [AutoText: Using fewer keystrokes in LibreOffice Writer  -  r/libreoffice](https://www.reddit.com/r/libreoffice/comments/51l82k/autotext_using_fewer_keystrokes_in_libreoffice/)
      * AutoText is LO Writer's way to allow shortcuts that are based on text.
      * It's similar to AutoCorrect, but AutoText can work with diagrams/drawings too!
    * ### [Adding bulk entries into autocorrect list - English  -  Ask LibreOffice](https://ask.libreoffice.org/t/adding-bulk-entries-into-autocorrect-list/39742)
      * [Export LibreOffice Autocorrect List - English  -  Ask LibreOffice](https://ask.libreoffice.org/t/export-autocorrect-list/6364/2)
    * ### [Using Regular Expressions in Text Searches  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/swriter/guide/search_regexp.html)
      * Keyphrases: CTRL+H, Find and Replace, Search and Replace, RegEx
    * ### [List of Regular Expressions  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/shared/01/02100001.html?&DbPAR=SHARED&System=WIN)
      * Contains examples of LO RegEx.
      * Is actually somewhat helpful.

  * ## Formatting
  * ### [Configure spacing between list elements - English  -  Ask LibreOffice](https://ask.libreoffice.org/t/configure-spacing-between-list-elements/15094/3)
  * ### [How to stop text body format bolding? - English  -  Ask LibreOffice](https://ask.libreoffice.org/t/how-to-stop-text-body-format-bolding/66305)
    * **Symptom:**
      * Upon highlighting two words (not even including any invisible newline characters), and then pressing `Ctrl+B` to bold it, it bolded the two highlighted words *and all the text below it* instead of just bolding the two words.
      * I had just created a new document, wrote stuff in it, and it automatically used the "Default Paragraph Style", so when I copied the text to a new document, Writer had preserved the style despite it not matching the document I copied the new document's contents to.
    * **Solution:** Change the Style type from something like "Default Paragraph Style" to "Normal", which is a non-header Style. Header styles assume that if you bold one character that's using that style, that you also want to bold every other (recursively) contiguous character that uses the same style.
  * ### [Ctrl+Shift in Writer switches to reversed-direction typing - English  -  Ask LibreOffice](https://ask.libreoffice.org/t/ctrl-rightshift-in-writer-switches-to-reverse-typing/133588)
    * **Symptom:** Text switched from being left-aligned to now being right-aligned, AND the arrow keys traverse the opposite direction as expected (left arrow key moves cursor rightward, right arrow key moves cursor leftward).
    * **Solution:** You can change from RTL (Right-to-Left) language configuration to LTR (Left-to-Right) language direction by pressing `Ctrl+Shift+A` **\***. The keybind that switches to RTL (non-Western language) is `Ctrl+Shift+D`, as found by `Tools` -> `Customize…` -> `Keyboard` -> Scrolling until I found that the keybinds tied to `Ctrl+Shift+A` and `Ctrl+Shift+D` were `Left-to-Right` and `Right-to-Left` respectively.
      * **\*** Any combination of `Shift` and `Ctrl`, whether the left key or right key (on your keyboard), will work.
        * E.g., `LCtrl+LShift+D`, `LCtrl+RShift+D`, `RCtrl+LShift+D`, and `RCtrl+RShift+D` all work the same.
  * ### Can't type code nor nested paragraphs into Writer document because it removes all indentation every time I press ENTER key
    * [Removing Indents When Pressing ENTER  -  r/libreoffice](https://www.reddit.com/r/libreoffice/comments/1pqep1h/removing_indents_when_pressing_enter/)
      * **Solution:** `Tools` > `AutoCorrect` > `AutoCorrect Options` > `Options` tab > 
        * Uncheck the option labeled `Delete spaces and tabs at beginning and end of paragraph` for BOTH `[M]` and `[T]`
        * Uncheck the option labeled `Delete spaces and tabs at end and start of line` for BOTH `[M]` and `[T]`
  * ### How to indent lines after copy-paste destroyed indentations?
    * **Solution:** `Find and Replace` -> Check/Mark/Fill the `Regular expressions` checkbox -> `Find:` `(^.)`, `Replace:` `\t$0` -> `Find Next` (click on an  area nearby to where you want to replace, and maneuver until the highlighted region of text is at the correct location) -> `Replace`
      * DO NOT do `replace all`, as this will indent EVERY line.
      * This does "find the beginning of a line and then the following character, and then replace the character with a TAB and the same character that was found"
  * ### How to switch the type of quotation marks and apostrophes from *directed* to *undirected*?
    * **Examples:**
      * Directed/Directional quotes: `“OK”`.  Undirected/Non-directional quotes: `"OK"`.
    * **Uses:**
      * Code written in LO Writer doesn't mess up (invalid character – expected non-directional quote)
      * Don't have to add an extra space before a word so that the direction is correct when entering an apostophe or quote, before deleting that extra space character
    * **Solution:**
      * `Tools` -> `AutoCorrect` (dropdown) -> `AutoCorrect Options` (item in dropdown) -> `Localized Options` (Tab) -> UNCHECK the boxes that say `Replace` that are above `Single Quotes` and `Double Quotes` -> Press `OK`.
        * The apostrophe and double quotes that are under the `Replace` boxes visually LOOK correct in the dialog box, but clicking on them shows that they are instead *directional* according to the symbol table!

  * ## Image
  * ### Image still takes up massive amount of visual space (cursor is very tall) after cropping it, when image is bullet-indented into page:
    * **Solution:**  Right click on the image. Go to `Align Objects`, then click any one of the three following options: `Base line at top`, `Base line centered`, `Base line at bottom`.
      * For some reason this works (successfully shortening the cursor height to where text positioning now makes sense). No idea why.
  * ### [How to apply crop of an image? - English  -  Ask LibreOffice](https://ask.libreoffice.org/t/how-to-apply-crop-of-an-image/133285)
    * **Not *exactly* mentioned in link:**  Writer does not overwrite the actual content of images that are pasted into the document. Writer only shows the cropped version of the image, leading to pretty large image sizes if you don't deal with them (lossy vs lossless format) BEFORE pasting them in.
      * This is also a data-privacy/cybersecurity issue. If you crop private data out of a sensitive image, the image can be uncropped, revealing the private data.



# Formula Editor
  * ### [Getting Started with LibreOffice Math](https://books.libreoffice.org/en/GS73/GS7309-GettingStartedWithMath.html)
  * ### [Entering Comments In Formula Editor  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/smath/guide/comment.html?DbPAR=MATH)
    * `a^2 = b^2  + c^2   %% This is a single-line comment`
  * ### [Format - Superscripts, Subscripts, newline, Stacks, Gaps/Spacing, Align - Formula Editor  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/smath/01/03090700.html?DbPAR=MATH)
  * ### [Manually Aligning Formula Parts  -  LibreOffice Help](https://help.libreoffice.org/latest/lo/text/smath/guide/align.html)
    * At the beginning of the line to display, use empty strings (`""`) and empty groups (`{}`), with help from `newline`, within the Formula Editor.
  * ### [Formula align on character - English  -  Ask LibreOffice](https://ask.libreoffice.org/t/formula-align-on-character/24119)
    * For alignment, you can use:
      * Inbuilt alignment:
        * `alignl{This section is left-aligned} but this section isn’t necessarily left-aligned`
        * `alignr{This section is right-aligned} but this section isn’t necessarily right-aligned`
        * `alignc{This section is center-aligned} but this section isn’t necessarily center-aligned`
      * Matrices: `matrix{r1c1 # r1c2~~~`` # r1c3   ##   r2c1~~` # r2c2 }`
      * Double double-quotes at the left (for left alignment only)
        * E.g., `""This formula text is left-aligned   newline This text is center-aligned, which is the default.`
      * Not mentioned in link:
        * Phantom: `phantom{invisible text that takes up space goes here}`
  * ### [Every type of bracket in Formula Editor  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/smath/01/03090500.html?DbPAR=MATH)
  * ### [Arrows, Dots, Other Math Symbols  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/smath/01/03091600.html?DbPAR=MATH)
  * ### [User-Defined Operators  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/smath/01/03090300.html?DbPAR=MATH)
    * I still can’t figure out how `oper` is supposed to work.
  * ### [Math Guide 25.2 - Chapter 2, Formulas in Writer](https://books.libreoffice.org/en/MG252/MG25202-FormulasInWriter.html)
    * **Explains:**
      * How to create numbered (indexed) formulas
        * E.g., "This is formula number 1 that I will be referencing later in the document"
      * How to change the WHOLE formula box’s positioning on the page (called “anchor”ing), whether to left or right or up or down, or inline with text as either a large or small formula, or standalone (i.e, NOT inline with text) as a large formula.
      * How to quickly enter a formula without creating a formula box that you need to manually enter and edit.

# Macros
  * ### [The Unofficial LibreOffice Basic Reference — LibreOffice Basic Reference 0.1 documentation](https://libreoffice-basic-reference.readthedocs.io/en/latest/)
  * ### [How to change formatting of multiple formulas simultaneously, instead of one-by-one - English  -  Ask LibreOffice](https://ask.libreoffice.org/t/how-to-change-multiple-formulas-simultaneously-instead-of-one-by-one/105924/21)
  * ### [Basic-Language Programming Examples  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/sbasic/guide/basic_examples.html)
  * ### [Basic-Language Development Tools  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/shared/guide/dev_tools.html?DbPAR=BASIC)
  * ### [Basic-Language Syntax  -  LibreOffice Help](https://help.libreoffice.org/latest/en-US/text/sbasic/shared/01020000.html?DbPAR=BASIC)

