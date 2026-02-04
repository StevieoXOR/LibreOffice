## Example of MathAutoCorrect

Before substituting key phrases:
<img alt="Before substituting key phrases" src="Assets/PreSubstitution.png" width=1000 height=600>
<br>

During substitution of key phrases
<img alt="During substitution of key phrases" src="Assets/DuringSubstitution.png" width=1000 height=600>
<br>

After substituting key phrases
<img alt="After substituting key phrases" src="Assets/AfterSubstitution.png" width=1000 height=600>
<br>
<br>

## Purpose of MathAutoCorrect

AutoCorrect doesn't exist when inside LibreOffice Math Formula Objects, so there is no
possibility of unofficial LO shortcuts (at least, if you do not want to have to arduously
click through the GUI to get to your user-defined formulas). This is especially annoying with long equations.

<br>

Also, if you forget certain patterns that LO Writer already uses, this macro lets
you simply write down the shortcut to some pre-defined rule, like `%idmat2`, that contains
the native LO Writer Math Formula pattern. Example native LO Writer Math Formula patterns:
* `left [` SomeContents `right ]`
* `left [` SomeContents `right none`
* `stack{` SomeContents `}`
* `matrix{ My_row1col1 # My_row1col2 ## My_row2 }`

<br>

This macro code lets you substitute keyphrases into their expanded form when the user is
inside (not merely selecting) the Math Formula Box Editor inside a LibreOffice Writer file, 
like converting (i.e., substituting)
* `%al ` into `%alpha`
* `%the ` into `%theta`
* `%sig ` into `%sigma`
* `%Sig ` into `%SIGMA`
* `%irt2` into `frac{1}{sqrt{2}}` (for "Inverse squareRooT of 2")
* `%mat2` into 
  ```
  left [
  matrix{
     a # b
  ## c # d
  }
  right ]
  ```
* `%idmat4` into
  ```
  left [
  matrix{
     1 # 0 # 0 # 0
  ## 0 # 1 # 0 # 0
  ## 0 # 0 # 1 # 0
  ## 0 # 0 # 0 # 1}
  right ]
  ```
* `%cases2` or `%piecewise2` or `%pw2` into
  ```
  left {
    stack{a, x>0 # b, x <= 0}
  right none
  ```
* `%cases4` or `%piecewise4` or `%pw4` into
  ```
  size*0{ phantom{Piecewise Function 4} }
  stack{%theta`=` # ` # `}
  size *3.75{\lbrace}
    stack{
      {x,```i>0}
    # {y,```i=0}
    # {z,```i<0}
    # {%alpha,`i notin setR}
    # {size *2.5{~}}
    }``````
  ```
* `%deriv` into `{{df} over {dx}}`
* Many more substitutions that have already been implemented.


## File Details

* `MathFormulaExpander.vb` is the file that contains the macro to run (**"Main_ExpandFormulaShortcuts"**)
  once you're inside a Formula Editor, as well as the macro that details a list of available substitutions
  (**"ListAvailableShortcuts"**) (which can be run either inside the main Writer document or inside the Formula Editor).
  * This is the file where you should add new rules (or modify old ones) to your liking.

* `MathFormulaExpander - GetFormulaObject_Experimenting.vb` is purely a development file that
contains attempts to allow substitution when the user's cursor has selected but not entered a math formula. 
All attempts so far have been unsuccessful.

* `MathFormulaExpander_ShortcutsTestbench.txt` is a file that should contain all the substitution
  rules that you can copy into a Formula Editor, then run the substitution macro to look for any
  unintended changes that would indicate you need to change the position or input string of a
  substitution rule.


## Notes

* This does not substitute the visual-only representation of the Math Formula.
  It substitutes the actual text inside the Math Formula (which then alters the visual representation).

* If desired, the Writer-recognized constants like %SIGMA and %sigma can be replaced
  with the actual symbols by modifying this macro and copying the actual unicode symbols
  online (or even from within Writer via the Symbols section).

* Adding new rules only needs to be done in one file, but that still sadly isn't as simple as the native AutoCorrect method. In other words, instead of opening a dialog box to add a (regular paragraph) word substitution (regular AutoCorrect), you must open and modify this Macro file (Math AutoCorrect - `MathFormulaExpander.vb`) (specifically the `ReplaceAllShortcuts` Function and `ListAvailableShortcuts` Sub).



### To Do

* Add shortcuts for sparse matrices, dot (vertical, horizontal, downright, downleft), %veps for %varepsilon.
* Add shortcuts for quantum gate matrix-representations (X,Y,Z,H,CX,CCX/Toffoli,SWAP, RX(theta),RY(theta),RZ(theta)), quantum state |i>, quantum |+> and |-> variants where fractions are separated.
* Improve this README to detail how to set up a keybind to auto-run the macro after pressing CTRL+SPACE,
  and link to a related macro & keybind tutorial.
* The To-Dos listed inside the Testbench file for automatically extracting the set of rules from
  `MathFormulaExpander.vb` and turning them into a Testbench file.
  * Hardcode the 1st and last ("%n")
    rules as a sentinel to search for when creating the Testbench file, which allows ignoring all
    the actual code with the help of fixed line widths in the rules section.
