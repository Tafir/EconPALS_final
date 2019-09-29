import pandas as pd
import sys
import numpy as np


def rename_sessions(df, rename=False):
    """
    :param df: a data frame containing the exported data from TopHat
    :param rename: boolean is true for when the function is called after original importing, e.g. when a new session is inserted
    :return: returns a df with all the 'lectures' renamed to 'session' with proper week and day labels
    """

    # declare counters: i for weekdays, j for weeks, k for semesters
    i = 1
    j = 2
    k = 1

    if rename is False:
        # replace "Lecture X" with more meaningful session names
        for session in df.columns:
            if 'Lecture' in session:
                # check if a week has passed, if so, increase the week counter
                if i == 5:

                    i = 1
                    j = j + 1

                    if j == 11:
                        j = 2
                        k += 1

                # declare a dictionary, assigning weekday strings to i values
                weekdays = \
                    {1: ('Tuesday s{}w{}'.format(k, j)),
                     2: ('Wednesday s{}w{}'.format(k, j)),
                     3: ('Thursday s{}w{}'.format(k, j)),
                     4: ('Friday 1 s{}w{}'.format(k, j))}

                # rename columns of the dataframe
                df = df.rename(index=str, columns={session: weekdays.get(i)})

                i += 1
        return df

    else:

        #df.columns = df.columns.str.replace('Friday 23 s1w3',
        #                                    'Friday 2 s1vv3')  # this is crap, drop it before release!!!
        for session in df.columns:
            if ('w' in session) and not ("Friday 2 " in session):
                # check if a week has passed, if so, increase the week counter
                if i == 5:

                    i = 1
                    j = j + 1

                    if j == 11:  # check if the semester is over, if so, increase the semester counter
                        j = 2
                        k += 1

                # declare a dictionary, assigning weekday strings to i values
                weekdays = \
                    {1: ('_Tuesday s{}w{}'.format(k, j)),
                     2: ('_Wednesday s{}w{}'.format(k, j)),
                     3: ('_Thursday s{}w{}'.format(k, j)),
                     4: ('_Friday 1 s{}w{}'.format(k, j))}

                # rename columns of the data frame
                df = df.rename(index=str, columns={session: weekdays.get(i)})
                i += 1

        return df


def drop_leaders_uuns(df, leaders_uuns):  # gets rid of double entries and leaders' emails
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :param leaders_uuns: a list containing leaders' uuns, which are to be excluded from the data frame
    :return: returns a corrected data frame
    """

    for i in df.index:
        uun = df.at[i, 'Username'][1:8]
        if uun in leaders_uuns:
            df = df.drop(index=i)
    return df


def correct_uuns(df):  # appends 'sms.' to UUN emails
    """

    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :return: a data frame with corrected uni emails
    """

    for i in df.index:
        email = df.at[i, 'Username']
        # check if it's a uni email, add 'sms.'
        if '@ed.ac.uk' in email:
            email_parts = email.split('@')
            email_parts[1] = '@sms.ed.ac.uk'
            email = ''.join(email_parts)
            df.at[i, 'Username'] = email
    return df


def read_file(filename, leaders_uuns=[''], columns_to_drop=[''], columns_to_merge=['']):  # creates a corrected and sorted data frame
    """
    :param filename: a string containing the file path to the TopHat export
    :param leaders_uuns: a list containing leaders' uuns, which are to be excluded from the data frame
    :param columns_to_drop: a list of session (lecture) labels which are to be excluded from the data frame
    :param columns_to_merge: a list of session (lecture) labels which are to be merged on import
    :return: a data frame of the exported data
    """

    df = pd.read_excel(io=filename, sheet_name='Attendance')

    if columns_to_drop != ['']:  # if there is a column to drop
        try:
            df = df.drop(columns=columns_to_drop)
        except KeyError:
            print(sys.exc_info()[1])  # outputs which key has not been found
            print("Column drop failed. Ignore if you are importing a formatted file")

    if columns_to_merge != ['']:
        try:
            merge_tuples = [(columns_to_merge[i], columns_to_merge[i+1]) for i in range(0, len(columns_to_merge), 2)]
            for pair in merge_tuples:
                df = merge_sessions(df, pair[0], pair[1])
        except:
            print(sys.exc_info()[1])

    if leaders_uuns != ['']:  # if the leaders uun field in the settings is not empty
        try:
            df = drop_leaders_uuns(df, leaders_uuns)
        except:
            print("Drop leaders UUNS failed. Ignore if you are importing a formatted file")
            print(e=sys.exc_info())

    df = df.sort_values(by='Username')
    df = rename_sessions(df)

    df = correct_uuns(df)
    return df


def studypals_read_file(filename, df, leaders_uuns):  # imports a StudyPALS excel file
    """
    :param filename: a string containing the file path to the StudyPALS (EconPALS1) TopHat export
    :param df: a data frame with the EconPALS TopHat data
    :param leaders_uuns: a list containing leaders' uuns, which are to be excluded from the data frame
    :return: a data frame of the exported data with the StudyPALS(EconPALS1) data appended to it
    """

    week = 2
    semester = 1
    temp_df = pd.read_excel(io=filename, sheet_name='Attendance')
    temp_df = temp_df.sort_values(by='Username')
    for session in temp_df.columns:
        if "Lecture" in session:
            if week == 11 and semester == 1:
                semester = 2
                week = 2
            temp_df = temp_df.rename(index=str, columns={session: 'Friday 2 s{}w{}'.format(semester, week)})
            week += 1
    for i in temp_df.index:
        uun = temp_df.at[i, 'Username'][1:8]
        if uun in leaders_uuns:
            temp_df = temp_df.drop(index=i)
    temp_df = temp_df.drop(
        columns=['Student ID', 'Email Address', 'First Name', 'Last Name', 'Attended', 'Excused', 'Absent'])
    temp_df = temp_df.rename(index=str, columns={'Username': 'Username_temp'})
    df = pd.concat([df, temp_df], axis=1, sort=False)
    df = df.fillna('A')
    for i in df.index:
        if df.at[i, 'Username'] == 'A' and df.at[i, 'Username_temp'] != 'A':
            df.at[i, 'Username'] = df.at[i, 'Username_temp']
    df = df.drop(columns=['Username_temp'])
    df = update_attendance(df)
    return df


def attendance_check(df, i, week, semester):  # checks whether a person i attended a session in a given week
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :param i: integer, index of a person being looked up
    :param week: integer, week label of the sessions being looked up
    :param semester: integer, semester label of the sessions being looked up
    :return: boolean, True if person i attended a session in a given week
    """

    if week != 0:
        if df.at[i, "Tuesday s{}w{}".format(semester, week)] == 'P':
            return True
        elif df.at[i, "Wednesday s{}w{}".format(semester, week)] == 'P':
            return True
        elif df.at[i, "Thursday s{}w{}".format(semester, week)] == 'P':
            return True
        elif df.at[i, "Friday 1 s{}w{}".format(semester, week)] == 'P':
            return True
        else:
            return False
    else:
        for week in range(2, 11):
            if df.at[i, "Tuesday s{}w{}".format(semester, week)] == 'P':
                return True
            elif df.at[i, "Wednesday s{}w{}".format(semester, week)] == 'P':
                return True
            elif df.at[i, "Thursday s{}w{}".format(semester, week)] == 'P':
                return True
            elif df.at[i, "Friday 1 s{}w{}".format(semester, week)] == 'P':
                return True
        return False


def get_emails(df, week=0, semester=1, to_file=False):  # produces a file with a mailing list with every attendee ever
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :param week: integer, week label of the sessions being looked up
    :param semester: integer, semester label of the sessions being looked up
    :param to_file: boolean, if True, writes to the export to a text file
    :return: a string with all the emails of people who attended sessions in a given week+semester
    """

    emails = ''  # create an empty string for stuff to be appended to

    # if you want to print to a file, for whichever reason
    if to_file:
        if week == 0:  # for all weeks
            text_file = open('mailing_list.txt', 'w')
            for i in df.index:
                if attendance_check(df, i, week, semester):
                    text_file.write(df.at[i, 'Username'])
                    text_file.write('\n')
                    emails += df.at[i, 'Username'] + '\n'
            text_file.close()
        else:
            text_file = open(('mailing_list_s{]w{}.txt'.format(semester, week)), 'w')
            for i in df.index:
                if attendance_check(df, i, week, semester):  # for a particular week
                    text_file.write(df.at[i, 'Username'])
                    text_file.write('\n')
                    emails += df.at[i, 'Username'] + '\n'
            text_file.close()

    # just creating a string with all the emails
    else:
        if week == 0:  # for all weeks
            for i in df.index:
                if attendance_check(df, i, week, semester):
                    emails += df.at[i, 'Username'] + '\n'
        else:
            for i in df.index:
                if attendance_check(df, i, week, semester):  # for a particular week
                    emails += df.at[i, 'Username'] + '\n'
    return emails


def regulars_list(df, number_of_sessions, to_file=False):
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :param number_of_sessions: integer, a cutoff number of sessions after which a person is considered a regular
    :param to_file: boolean, if True, writes to a file
    :return: a string with all the regulars' emails
    """

    regulars = ''
    if to_file:
        text_file = open('regulars_list.txt', 'w')
        for i in df.index:
            if df.at[i, 'Attended'] > number_of_sessions:
                text_file.write(df.at[i, 'Username'])
                text_file.write('\n')
                regulars += df.at[i, 'Username'] + '\n'
        text_file.close()
    else:
        for i in df.index:
            if df.at[i, 'Attended'] > number_of_sessions:
                regulars += df.at[i, 'Username'] + '\n'
    return regulars


def merge_sessions(df, session1, session2):
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :param session1: string, 1st session to be merged, column will be preserved
    :param session2: string, 2nd session to be merged, column will be dropped
    :return: a data frame with the sessions merged
    """

    for i in df.index:
        if df.at[i, session1] == 'P':
            df.at[i, session2] = 'P'
    df = df.drop(columns=session2)
    return df


def attendance_count(df):
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :return: a list of tuples, first value is session name the second is the session's attendance count
    """

    attendance = []
    for column in df.columns:
        if 'day' in column:
            attendance.append((column, (len(df.index) - df[column].value_counts().values[0])))
    return attendance


# updates attendance after importing the StudyPALS file
def update_attendance(df):
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :return: a data frame with the StudyPALS (EconPALS1) attendance properly accounted for
    """

    for i in df.index:
        count = 0
        for column in df.columns:
            if df.at[i, column] == 'P':
                count += 1
        df.at[i, 'Attended'] = count
    return df


def get_attendance(df, email):
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :param email: string, an email of a person being looked up
    :return: string, a list of all the sessions a person has attended
    """
    attendance_string = ''

    # if the user inputs something looking like a student number, add @sms.ed.ac.uk to it
    if "@" not in email and email[0] == 's':
        email += "@sms.ed.ac.uk"

    if any(df['Username'].isin([email])):
        counter = 0

        # generates a 1 row df with only the person's attendance data in it
        temp_df = df[df["Username"] == email]

        for session in temp_df:
            if temp_df[session].values[0] == 'P':
                counter += 1
                attendance_string += session + "\n"
        return "{} attended a total of {} sessions: \n{}".format(email, counter, attendance_string)
    else:
        return "Invalid username entered"


if __name__ == "__main__":
    raise NotImplementedError("Run EconPALS.py to launch the app")