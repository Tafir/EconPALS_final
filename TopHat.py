import pandas as pd
import sys
import numpy as np

def drop_leaders_uuns(df, leaders_uuns):  # gets rid of double entries and leaders' emails
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :param leaders_uuns: a list containing leaders' uuns, which are to be excluded from the data frame
    :return: returns a corrected data frame
    """

    for i in df.index:
        uun = df.at[i, 'Email'][1:8]
        if uun in leaders_uuns:
            df = df.drop(index=i)
    return df



def correct_new_format(dataframe):
    """

    :param dataframe: the data frame containing raw excel input
    :return: a data frame containing a proper attendance sheet
    """
    columns = [c for c in dataframe.columns if "weight" not in c.lower()]
    df_clean = dataframe[columns]

    attendance_columns = columns[7:]
    df_attendance = pd.DataFrame()
    for column in df_clean:
        if column in attendance_columns:
            df_attendance[column] = pd.to_numeric(df_clean[column], errors="coerce")

    # Not the cleanest way, rewrite later
    df_attendance[df_attendance>=0] = "P"
    df_attendance = df_attendance.fillna("A")
    df_clean[attendance_columns] = df_attendance

    df_clean = df_clean.drop(df_clean[df_clean['Student Name']=='Class Average'].index)
    # FIXME: This is to overcome some indexing issues for attendance checking. There must be a neater way
    df_clean = df_clean.reset_index()
    return df_clean


def read_file(filename, leaders_uuns=[''], columns_to_drop=[''], columns_to_merge=['']):  # creates a corrected and sorted data frame
    """
    :param filename: a string containing the file path to the TopHat export
    :param leaders_uuns: a list containing leaders' uuns, which are to be excluded from the data frame
    :param columns_to_drop: a list of session (lecture) labels which are to be excluded from the data frame
    :param columns_to_merge: a list of session (lecture) labels which are to be merged on import
    :return: a data frame of the exported data
    """

    if ".csv" in filename:
        df_raw = pd.read_csv(filename)
    else:
        df_raw = pd.read_excel(io=filename)

    df = correct_new_format(df_raw)

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
            print(sys.exc_info())

    df = df.sort_values(by='Email')

    return df


def attendance_check(df, i, week, semester):  # checks whether a person i attended a session in a given week
    """
    :param df: a data frame containing the exported data from TopHat, with lectures renamed
    :param i: integer, index of a person being looked up
    :param week: integer, week label of the sessions being looked up
    :param semester: integer, semester label of the sessions being looked up
    :return: boolean, True if person i attended a session in a given week
    """
    try:
        if week == 0:
            # FIXME: This is ugly. Probably pass week as empty string by default?
            week = ''
 
        # Get all columns for this person
        row = df.iloc[i]

        # Separate attendance related columns for a given semester-week combination. If week=0, pull entire semester
        attendance_columns = [column for column in row.index if f"S{semester}W{week}" in column]

        # Check if value of any of the attendance columns is equal to P
        return row[attendance_columns].isin(['P']).any()

    except:
        print("An error has occured during attendance checking")
        print(sys.exc_info())
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
                    text_file.write(df.at[i, 'Email'])
                    text_file.write('\n')
                    emails += df.at[i, 'Email'] + '\n'
            text_file.close()
        else:
            text_file = open(('mailing_list_s{}w{}.txt'.format(semester, week)), 'w')
            for i in df.index:
                if attendance_check(df, i, week, semester):  # for a particular week
                    text_file.write(df.at[i, 'Email'])
                    text_file.write('\n')
                    emails += df.at[i, 'Email'] + '\n'
            text_file.close()

    # just creating a string with all the emails
    else:
        for i in df.index:
            if attendance_check(df, i, week, semester):  # for a particular week
                emails += df.at[i, 'Email'] + '\n'
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
            if df.at[i, 'Attendance'] > number_of_sessions:
                text_file.write(df.at[i, 'Email'])
                text_file.write('\n')
                regulars += df.at[i, 'Email'] + '\n'
        text_file.close()
    else:
        for i in df.index:
            sessions_attended = df.at[i, 'Attendance']
            try:
                if int(df.at[i, 'Attendance']) >= number_of_sessions:
                    regulars += df.at[i, 'Email'] + '\n'
            except:
                pass
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

    if any(df['Email'].isin([email])):
        counter = 0

        # generates a 1 row df with only the person's attendance data in it
        temp_df = df[df["Email"] == email]

        for session in temp_df:
            if temp_df[session].values[0] == 'P':
                counter += 1
                attendance_string += session + "\n"
        return "{} attended a total of {} sessions: \n{}".format(email, counter, attendance_string)
    else:
        return "Invalid Email entered"


if __name__ == "__main__":
    raise NotImplementedError("Run EconPALS.py to launch the app")