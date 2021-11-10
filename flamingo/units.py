import re

def FeetInchToFloat(lengthString):
    footInchPattern = re.compile(
        r"^\s*(?P<minus>-)?\s*("
        r"(((?P<feet>[\d.]+)')[\s-]*"
        r"((?P<inch>[\d.]+)?[\s-]*"
        r"((?P<numer>\d+)/(?P<denom>\d+))?\"?)?)|"
        r"(((?P<inch>[\d.]+)?[\s-]*"
        r"((?P<numer>\d+)/(?P<denom>\d+))?\")?)|"
        r"((?P<feet>[\d.]+)([\s-]+(?P<inch>[\d.]+))?"
        r"([\s-]+(?P<numer>\d+)/(?P<denom>\d+))?)" #if only spaces are entered
        r")\s*$"
    )
    match = footInchPattern.search(lengthString)
    lengthFloat = None
    if match:
        matches = match.groupdict()
        for key, value in matches.items():
            if key == "minus":
                if value == "-":
                    matches[key] = -1
                else:
                    matches[key] = 1
            if key in ["feet", "inch", "numer", "denom"]:
                if value is None:
                    if key == "denom":
                        matches[key] = 1
                    else:
                        matches[key] = 0
                else:
                    matches[key] = float(value)
        lengthFloat = (
            matches["feet"] + (
                (matches["inch"] + (matches["numer"]/matches["denom"])) / 12
            )
        ) * matches["minus"]
    return lengthFloat

def FloatToFeetInchString(lengthFloat):
    feet = int(lengthFloat)
    lengthInches = abs(lengthFloat * 12)
    inches = lengthInches % 12
    remainder = round(inches, 5) - int(inches)
    remainderCheckCount = 0
    while remainder != 0 and remainderCheckCount < 8:
        remainderCheckCount += 1
        denominator = pow(2, remainderCheckCount)
        lengthFraction = lengthInches * denominator
        numerator = lengthFraction % denominator
        remainder = round(numerator, 5) - int(numerator)
    if numerator and numerator > 0:
        fraction = " {}/{}".format(int(numerator), int(denominator))
    else:
        fraction = ""
    lengthString = "{}'-{}{}\"".format(feet, int(inches), fraction)
    return lengthString