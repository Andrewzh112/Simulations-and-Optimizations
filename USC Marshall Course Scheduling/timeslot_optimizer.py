from gurobipy import Model, GRB
import os
import re

def optimize(inputFile,outputFile,classes_same_time=60,w_z=0.1,w_f=0.5,w_s=0.5,w_w=0.5,install_pandas=False):
    """
    optimizes for the department timeslot assignment fairness

    Args:
        inputFile (str): The input data file name
        outputFile (str): The output data file name, also include .xlsx extension
        classes_same_time (int): (Optional) The maximum number of classes that can be assigned to one time slot (Default=60)
        w_z (float): (Optional) The weight of faculty load factor in objective (Default=0.1)
        w_f (float): (Optional) The weight of faculty importance factor in objective (Default=0.5)
        w_s (float): (Optional) The weight of student importance factor in objective (Default=0.5)
        w_w (float): (Optional) The weight assigned to day of week balance in objective (Default=0.5)
        install_pandas (bool): (Optional) Whether the user needs to install pandas to use this tool (Default=False)

    Returns:
        None

    Output File:
        Optimized Solution.xlsx
    """
    from gurobipy import Model, GRB
    import os
    import re
    # if user does not have pandas installed:
    if install_pandas:
        os.system('cmd /c "pip install pandas"')
    import pandas as pd

    # load in sheets
    faculty = pd.read_excel(inputFile, sheet_name='Faculty Satisfaction', index_col=0)
    student = pd.read_excel(inputFile, sheet_name='Student Satisfaction')
    course = pd.read_excel(inputFile, sheet_name='Course Information', index_col=0).dropna()
    course['Student_Type'] = ['Graduate student' if int(i[1][0]) >= 5 else 'Undergraduate student'
                              for i in course['cour_sec'].str.split('-')]
    FS = course[['cour_sec', 'first_instructor']].merge(faculty, left_on='first_instructor',
                                                        right_index=True, how='left') \
                                                        .drop(columns='first_instructor') \
                                                        .set_index('cour_sec')
    student_score = course[['cour_sec', 'department', 'Student_Type']].merge(student,
                                        left_on=['department', 'Student_Type'], \
                                        right_on=['Dept', 'Student_Type'],
                                        how='left') \
                                        .drop(columns=['department', 'Dept', 'Student_Type', 'Long_Dept']) \
                                        .set_index('cour_sec')
    student_score.fillna(student_score.mean(), inplace=True)

    # manipulation and subset
    l = []
    for i in range(5):
        l.append(student_score)
    SS = pd.concat(l, axis=1)
    SS.columns = faculty.columns
    D = pd.DataFrame(0, index=course['cour_sec'], columns=course['department'].unique())
    for k in course['department'].unique():
        D[k] = [int(i) for i in (course['department'] == k)]

    # data variables, refer to formulation notebook for data variable meanings
    I = course['cour_sec']
    J = faculty.columns
    J1 = J[:28]
    J2 = J[-28:]
    J3 = [j for j in J2 if 'F' in j]
    K = course['department'].unique()
    F = faculty.index
    n = course['department'].value_counts()
    M = []
    morning_cutoff = 18
    for j in J:
        if int(re.search(r'(\d+)', j)[0]) < morning_cutoff:
            M.append(j)
    s = course.set_index('cour_sec')['Should be night']
    profile_codes = course.set_index('cour_sec').groupby('Profile code')['Profile code']
    G = [profile[1].index.tolist() for profile in profile_codes]
    W = [[j for j in J if 'M' in j],
         [j for j in J if 'T' in j],
         [j for j in J if 'W' in j],
         [j for j in J if 'H' in j],
         [j for j in J if 'F' in j]]
    J4 = []
    prev = None
    for w in W:
        if prev is None:
            prev = w[-1]
            continue
        J4.append([prev, w[0]])
        prev = w[-1]

    # define model, decision variables and set the objective for optimization
    mod = Model()
    x = mod.addVars(I, J, vtype=GRB.BINARY)
    up = mod.addVar(lb=0)
    low = mod.addVar(lb=0)
    Z = mod.addVar(lb=0, vtype=GRB.INTEGER)
    WWu = mod.addVar(lb=0, vtype=GRB.INTEGER)
    WWl = mod.addVar(lb=0, vtype=GRB.INTEGER)
    mod.setObjective(up - low + w_z * Z + w_w * (WWu - WWl), sense=GRB.MINIMIZE)

    # Department Balance
    exp = {}
    for k in K:
        nk = n[k]
        exp[k] = sum(
            (w_f * FS.loc[i, j] * D.loc[i, k] * x[i, j] + w_s * SS.loc[i, j] * D.loc[i, k] * x[i, j]) for i in I for j
            in J)
        mod.addConstr(exp[k] / nk <= up)
        mod.addConstr(exp[k] / nk >= low)

    # Faculty Time Conflict
    for f in F:
        for j in J:
            I_f = course[course.first_instructor == f]['cour_sec']
            mod.addConstr(sum(x[i, j] for i in I_f) <= 1)

    # Faculty Blackout
    for i in I:
        for j in J:
            mod.addConstr(x[i, j] <= FS.loc[i, j])

    # Unit Constraint
    for i in I:
        mod.addConstr(sum(x[i, j] for j in J) == course.loc[course.cour_sec == i, 'units'])

    # Classroom Constraint
    for j in J:
        mod.addConstr(sum(x[i, j] for i in I) <= classes_same_time)

    # Back to Back for Two Units Class
    for i in course[course.units == 2]['cour_sec']:
        num = 0
        for j1 in J[:(len(J) - 2)]:
            for j2 in J[num + 2:]:
                mod.addConstr(x[i, j1] + x[i, j2] <= 1)
            num += 1

    # Back to Back for Four Units Class
    for i in course[course.units == 4]['cour_sec']:
        num = 0
        for j1 in J1[:(len(J1) - 2)]:
            for j2 in J1[num + 2:]:
                mod.addConstr(x[i, j1] + x[i, j2] <= 1)
            num += 1

    # For each back to back sessions, they should be on the same day
    for i in I:
        for j4 in J4:
            mod.addConstr(x[i, j4[0]] + x[i, j4[1]] <= 1)

    # Same Time Session for Four Units Class
    for i in course[course.units == 4]['cour_sec']:
        num = 0
        for j in J1:
            mod.addConstr(x[i, j] == x[i, J[num + 28]])
            num += 1

    # No Class on Friday for Four Units Class
    for i in course[course.units == 4]['cour_sec']:
        for j in J3:
            mod.addConstr(x[i, j] == 0)

    # Night class constraint
    for i in I:
        for j in M:
            mod.addConstr(x[i, j] <= (1 - s.loc[i]))

    # Classes that are for same student profiles should not happen at same time
    for g in G:
        for j in J:
            mod.addConstr(sum(x[i, j] for i in g) <= 1)

    # Faculty load
    for w in W:
        for f in F:
            I_f = course[course.first_instructor == f]['cour_sec']
            mod.addConstr(sum(x[i, j] for j in w for i in I_f) <= Z)

    # day of week time load balance (except fridays)
    for w in W:
        if 'F' in w[0]:
            continue
        mod.addConstr(sum(x[i, j] for j in w for i in I) >= WWl)
        mod.addConstr(sum(x[i, j] for j in w for i in I) <= WWu)

    # Optimize
    mod.setParam('outputflag',False)
    mod.optimize()
    
    # Summary sheet for outputfile, including all department scores
    faculty_scores, student_scores = {}, {}
    for k in course['department'].unique():
        faculty_scores[k] = sum(FS.loc[i, j] * D.loc[i, k] * x[i, j].x for i in I for j in J) / n.loc[k]
        student_scores[k] = sum(SS.loc[i, j] * D.loc[i, k] * x[i, j].x for i in I for j in J) / n.loc[k]
    deps = [dep for dep in course['department'].unique()] + ['Max hour for any given day for every professor']
    scores = [round(exp[dep].getValue() / n.loc[dep], 2) for dep in course['department'].unique()] + [Z.x]
    summary_sheet = pd.DataFrame(scores, index=deps, columns=['Mean Score'])
    summary_sheet['Faculty Mean Score'] = list(faculty_scores.values()) + ['']
    summary_sheet['Student Mean Score'] = list(student_scores.values()) + ['']

    # Assigned course times sheet
    courses_times = [j for i in I for j in J if x[i, j].x]
    courses_code = [i for i in I for j in J if x[i, j].x]
    course_sheet = pd.DataFrame(courses_times, index=courses_code, columns=['Assigned Time'])
    
    # Assigned course times by day of week
    time_per_day = [sum(x[i, j].x for j in w for i in I) for w in W]
    day_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    dow_sheet = pd.DataFrame(time_per_day, index=day_of_week, columns=['Units Assigned'])

    # Master timetable by department
    master_schedule = {}
    dow_mapper = {day[0]: day for day in day_of_week}
    dow_mapper['H'] = 'Thursday'
    dow_mapper['T'] = 'Tuesday'
    for dep in course['department'].unique():
        master_schedule[dep] = pd.DataFrame(index=range(8, 22), columns=day_of_week)
        for course_id, coursetime in course_sheet.iterrows():
            if D[dep].loc[course_id]:
                time = coursetime.iloc[0]
                if type(master_schedule[dep].loc[int(time[1:]), dow_mapper[time[0]]]) == float and pd.isna(
                        master_schedule[dep].loc[int(time[1:]), dow_mapper[time[0]]]):
                    master_schedule[dep].loc[int(time[1:]), dow_mapper[time[0]]] = [course_id]
                else:
                    master_schedule[dep].loc[int(time[1:]), dow_mapper[time[0]]].append(course_id)

    # Write all output sheets to output file
    writer = pd.ExcelWriter(outputFile)
    summary_sheet.to_excel(writer, sheet_name='Summary Sheet')
    course_sheet.to_excel(writer, sheet_name='Assigned Times')
    dow_sheet.to_excel(writer, sheet_name='Day of Week Assignment Stats')
    for dep in sorted(course['department'].unique()):
        master_schedule[dep].to_excel(writer, sheet_name=f'{dep} Schedule')
    writer.save()

if __name__=='__main__':
    import sys
    # Arguments going into opimization function in order
    args=['timeslot_optimizer.py',
          'inputFile',
          'outputFile',
          'classes_same_time',
          'w_z',
          'w_f',
          'w_s',
          'w_w',
          'install_pandas']

    # Check if missing or too many arguments
    if len(sys.argv)<3 or len(sys.argv)>9:
        print(f'Correct syntax: python {" ".join(args)}')

    # Save arguments to dictionary
    else:
        inputFile=sys.argv[1]
        outputFile=sys.argv[2]
        default_args={'classes_same_time':60,
                      'w_z':0.1,
                      'w_f':0.5,
                      'w_s':0.5,
                      'w_w': 0.5,
                      'install_pandas':False}

        # Process arguments to appropriate variable type
        if len(sys.argv) > 3:
            for i in range(len(sys.argv)):
                if i < 3:
                    continue
                if i != 8:
                    default_args[args[i]] = float(sys.argv[i])
                    assert (type(default_args[args[i]]) == float)
                if i == 8:
                    default_args[args[i]] = bool(sys.argv[i])
                    assert(type(default_args[args[i]])==bool)

        # If input file exists, execute optimize function
        if os.path.exists(inputFile):
            optimize(inputFile=inputFile,
                     outputFile=outputFile,
                     classes_same_time=(default_args['classes_same_time']),
                     w_z=default_args['w_z'],
                     w_f=default_args['w_f'],
                     w_s=default_args['w_s'],
                     w_w=default_args['w_w'],
                     install_pandas=default_args['install_pandas'])
            print(f'Successfully optimized. Results in "{outputFile}"')
        # File not found exception
        else:
            print(f'File "{inputFile}" not found!')

