from datetime import datetime
import tempfile
import shutil
import re
import codecs

def TimestampAsDateTime(string):
    return datetime.strptime(
        string[10:29],
        format="%Y-%m-%d %H:%M:%S"
    )

def OutputMD(output, string, printReport=True):
    if printReport:
        output.print_md(string)

def GetWorksharingReport(slogPath, output, printReport):
    """Print status of worksharing sessions for the specified sLog file

    Args:
        slogPath (str): path to the revit worksharing log file
        output (pyRevit.script.output): pyRevit object for output
        printReport (bool): True if report is to be printed

    Returns:
        [type]: [description]
    """
    fd, tempPath = tempfile.mkstemp()
    shutil.copy2(slogPath, tempPath)

    sessions = {}
    currentSessionId = None
    minStart = datetime.max
    maxEnd = datetime.now()
    with codecs.open(tempPath, "r", encoding="UTF-16") as f:
        for line in f:
            if line.startswith(" ") and currentSessionId:
                m = re.match(r" (.+)=.*\"(.*)\"", line)
                if m:
                    sessions[currentSessionId][m.group(1)] = m.group(2)
            elif re.findall(r">Session", line):
                currentSessionId = line[0:9]
                startDateTime = TimestampAsDateTime(line)
                sessions[currentSessionId] = {"start": startDateTime}
                minStart = startDateTime if startDateTime < minStart \
                    else minStart
            elif re.findall(r"<Session", line) and currentSessionId:
                currentSessionId = line[0:9]
                endDateTime = TimestampAsDateTime(line)
                sessions[currentSessionId]['end'] = endDateTime
            elif re.findall(r"<STC[^:]", line):
                sessions[currentSessionId]['lastSync'] = TimestampAsDateTime(
                    line
                )
            elif line[0:9] in sessions:
                sessions[currentSessionId]['lastActive'] = TimestampAsDateTime(
                    line
                )

    for session in sessions:
        if "end" in sessions[session]:
            endDateTime = sessions[session]['end']
        else:
            endDateTime = datetime.now()
            if "lastSync" in sessions[session]:
                sessions[session]['timeSinceLastSync'] = \
                    datetime.now() - sessions[session]['lastSync']
        sessions[session]['sessionLength'] = \
            endDateTime - sessions[session]['start']

    activeSessions = [
        session for session, sessionInfo in sessions.items()
        if "end" not in sessionInfo
    ]
    sortedSessionKeys = sorted(
        sessions,
        key=lambda x: sessions[x]['start']
    )
    OutputMD(output, "# Worksharing Log Info", printReport)
    OutputMD(output, "Active: {}".format(maxEnd - minStart), printReport)
    OutputMD(output, "central: {}".format(list(sessions.values())[0]['central']), printReport)
    OutputMD(output, "Total Sessions: {}".format(len(sessions)), printReport)
    OutputMD(output, "Active Sessions: {}".format(len(activeSessions)), printReport)
    OutputMD(output, "", printReport)
    OutputMD(output, "# Sessions", printReport)
    keysToPrint = {
        "build": "Revit Version",
        "host": "Computer Name",
        "end": "Session End",
        "sessionLength": "Session Length",
        "lastActive": "Last Activity",
        "lastSync": "Last Sync To Central",
        "timeSinceLastSync": "Time Since Last Sync",
    }
    
    for session in sortedSessionKeys:
        values = sessions[session]
        icon = ":white_heavy_check_mark:" if "end" in values else ":cross_mark:"
        output.print_md(
            "## {}: {} {}".format(values['user'], values['start'], icon)
        )
        OutputMD(output, "id: {}".format(session), printReport)
        for key, title in keysToPrint.items():
            if key in values:
                OutputMD(output, "{}: {}".format(title, values[key]), printReport)
        OutputMD(output, "", printReport)

    return sessions
