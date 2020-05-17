def optimize(inputFile, outputFile):
    import pandas as pd
    prefs=pd.read_excel(inputFile,header=[0,1,2],sheet_name='Preferences',index_col=0)
    names=prefs.index
    shifts=prefs.columns
    shift_id=shifts.get_level_values(2)
    prefs.columns=shift_id
    requirements = pd.read_excel(inputFile, sheet_name='Requirements',index_col=0)
    
    S = prefs.columns
    I = prefs.index
    N = int((len(shift_id)) / 21)
    Nights = [n for n in range(2, (len(shift_id)-1), 3)]
    l = shift_id[-1]
    A1 = 10
    q = 6
    dshift = 21
    
    from gurobipy import Model, GRB
    mod = Model()
    x = mod.addVars(I,S,vtype=GRB.BINARY)
    Lshift = mod.addVar(vtype=GRB.INTEGER)
    Lnight = mod.addVar(vtype=GRB.INTEGER)
    Ushift = mod.addVar(vtype=GRB.INTEGER)
    Unight = mod.addVar(vtype=GRB.INTEGER)
    mod.setObjective(sum(prefs.loc[i,s]*x[i,s] for i in I for s in S) - 100*(Ushift-Lshift) - 150*(Unight-Lnight),sense=GRB.MAXIMIZE)
    for s in S:
        mod.addConstr(sum(x[i,s] for i in I) == requirements.loc[s,'persons'])
    for n in range(N):
        for i in I:
            mod.addConstr(sum(x[i,n*dshift+s] for s in range(dshift))<=q)
    for i in I:
        for s in range(1,l+1):
            mod.addConstr(x[i,s-1]+x[i,s]<=1)
    for i in I:
        for s in Nights:
            mod.addConstr(x[i,s-2]+x[i,s-1]+x[i,s+2]+x[i,s+1]<=A1*(1-x[i,s]))
    for i in I:
        mod.addConstr(x[i,l-2]+x[i,l-1]<=A1*(1-x[i,l]))
    for i in I:
        for s in S:
            mod.addConstr(x[i,s]<=prefs.loc[i,s])
    for i in I:
        mod.addConstr(Lshift<=sum(x[i,s] for s in S))
        mod.addConstr(sum(x[i,s] for s in S)<=Ushift)
    for i in I:
        mod.addConstr(Lnight<=sum(x[i,s] for s in Nights+[l]))
        mod.addConstr(sum(x[i,s] for s in Nights+[l])<=Unight)
    mod.setParam('OutputFlag',False)
    mod.optimize()
    
    schedule=pd.DataFrame('',index=names,columns=shift_id)
    for i in I:
        for s in S:
            if x[i,s].x:
                schedule.loc[i,s] = x[i,s].x
    schedule.columns=shifts

    summary=pd.Series(name='Value')
    summary['Objective']=mod.objval
    summary['Total preference score']=sum(prefs.loc[i,s]*x[i,s].x for i in I for s in S)
    summary['Shift inequality']=Ushift.x-Lshift.x
    summary['Night inequality']=Unight.x-Lnight.x
    summary

    writer=pd.ExcelWriter(outputFile,datetime_format='m/dd')
    schedule.to_excel(writer,sheet_name='Schedule')
    summary.to_excel(writer,sheet_name='Summary')
    writer.save()
    
    return

if __name__=='__main__':
    import sys, os
    if len(sys.argv)!=3:
        print('Correct syntax: python optimize.py inputFile outputFile')
    else:
        inputFile=sys.argv[1]
        outputFile=sys.argv[2]
        if os.path.exists(inputFile):
            optimize(inputFile,outputFile)
            print(f'Successfully optimized. Results in "{outputFile}"')
        else:
            print(f'File "{inputFile}" not found!')