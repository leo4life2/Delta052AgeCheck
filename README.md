# Delta052AgeCheck

This is a simple Python script that checks the ages of vaccination patients for the Delta052 spreadsheet.

All constants (Age limitations for shots) are in the `ShotType` class at top of the script.

If any limitation needs to be changed, just change the number inside the array, and everything else should work fine.

If you need to add a new vaccine type, just create a new class under `ShotType` and add in the class attributes in the same format as other classes. Then add the string you want to detect for in the array at the bottom of `ShotType`. Lastly, add an entry into the dictionary with the string as key and the new vaccine's class as the value.
