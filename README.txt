

This Addin is a utility for Revit that allows the user to "cleanup" Text object Leaders by setting the first segment of a leader line to a specific angle.  It moves the elbow point of a leader based on the specified angle.   The driving concept was to easily implement a drafting style for details in which all leader lines are straight and them bend at the same angle.  This is tedious to do by hand, but this utility makes easy to do.

Usage:
First: Set the desired angle from the list in the Ribbon.  
Second: Click "Adjust Leaders" and start selecting the Text objects to adjust.  You will be able to select text objects until you exit the utility.  
Last: Hit "Escape" to exit the utility.

The Rules (Limitations):
Due to what is available in the c# API for Revit, only the end of a leader, and the elbow point are available for accurate manipulation.  This means that it is up to the User to make sure the first segement is horizontal.  The second segment wil then be bent at the specified angle without moving the Arrowhead.  Only Text objects are currently supported.


Angles:
Currently only 3 angles are supported: 60, 45, & 30.  This may change in the future, but currently this can only be changed in the source coode.