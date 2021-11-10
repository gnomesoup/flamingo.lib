from pyrevit import script
from pyrevit import forms
import re

def AskForLength(defaultLengthFloat):
    lengthFloat = None
    whileCount = 0
    while lengthFloat <= 0:
        if whileCount > 3:
            script.exit()
        whileCount = whileCount + 1
        lengthString = forms.ask_for_string(
            default=str(defaultLengthFloat),
            title="Enter interval spacing in feet & inches"
        )
        if lengthString is None:
            return
        footInchPattern = re.compile(
            r'(-?\d+(\.\d+)?)\'?( - | |-)?(\d+(\.\d+)?)? ?(\d+)?\/?(\d+)?'
        )
        match = footInchPattern.match(lengthString)
        if match is None:
            forms.alert("Invalid length entry. Please try again.")
        else:
            lengthFloat = 0
            if match.group(6) and match.group(7):
                lengthFloat = lengthFloat + (float(match.group(6)) / \
                    float(match.group(7)) / 12)
            if match.group(4):
                lengthFloat = lengthFloat + (float(match.group(4)) / 12)
            if match.group(1):
                lengthFloat = lengthFloat + float(match.group(1))
        if lengthFloat <= 0:
            print(
                "We need a bigger size than " +
                str(lengthFloat) +
                ". Please try again."
                )
    return lengthFloat
