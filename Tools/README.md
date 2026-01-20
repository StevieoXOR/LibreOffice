## Purpose
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

## Notes

* This does not substitute the visual-only representation of the Math Formula.
  It substitutes the actual text inside the Math Formula (which then alters the visual representation).

* If desired, the Writer-recognized constants like %SIGMA and %sigma can be replaced
  with the actual symbols by modifying this macro and copying the actual unicode symbols
  online (or even from within Writer via the Symbols section).


## File Details

* `MathFormulaExpander.vb` is the file that contains the macro to run (**"Main_ExpandFormulaShortcuts"**)
  once you're inside a Formula Editor, as well as the macro that details a list of available substitutions
  (**"ListAvailableShortcuts"**).
  * This is the file where you should add new rules (or modify old ones) to your liking.

* `MathFormulaExpander - GetFormulaObject_Experimenting.vb` is purely a development file that
contains attempts to allow substitution when the user's cursor has selected but not entered a math formula. 
All attempts so far have been unsuccessful.

* `MathFormulaExpander_ShortcutsTestbench.txt` is a file that should contain all the substitution
  rules that you can copy into a Formula Editor, then run the substitution macro to look for any
  unintended changes that would indicate you need to change the position or input string of a
  substitution rule.


### To Do

* Add quantum equation and qubit shorthands.
* Improve this README to detail how to set up a keybind to auto-run the macro after pressing CTRL+SPACE,
  and link to a related macro & keybind tutorial.
* The To-Dos listed inside the Testbench file for automatically extracting the set of rules from
  `MathFormulaExpander.vb` and turning them into a Testbench file.
  * Hardcode the 1st and last ("%n")
    rules as a sentinel to search for when creating the Testbench file, which allows ignoring all
    the actual code with the help of fixed line widths in the rules section.
