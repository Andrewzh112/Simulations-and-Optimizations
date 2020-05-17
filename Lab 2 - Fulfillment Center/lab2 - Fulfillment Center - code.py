def optimize(inputFile, outputFile):
    import pandas as pd
    import gurobipy as grb
    import numpy as np
    fc = pd.read_excel(inputFile, sheet_name='Fulfilment Centers',index_col=0)
    region = pd.read_excel(inputFile, sheet_name='Regions',index_col=0)
    distances = pd.read_excel(inputFile, sheet_name='Distances',index_col=0)
    items = pd.read_excel(inputFile, sheet_name='Items',index_col=0)
    demand = pd.read_excel(inputFile, sheet_name='Demand',index_col=0)

    mod=grb.Model()

    # Preparing data
    I=fc.index
    J=region.index
    K=items.index

    # Defining decision variables
    n=mod.addVars(I, J, K)

    # Defining objective and constraints
    mod.setObjective(sum(1.38*items.loc[k,'shipping_weight']*\
                         distances.loc[j,i]*n[i,j,k] for k in K for j in J for i in I))

    for j in J:
        for k in K:
            mod.addConstr(sum(n[i,j,k] for i in I)>=demand.loc[k,j])
    
    #saving capacity constraints for shadow price
    capacity_dict = {}
    for i in I:
        capacity_dict[i] = mod.addConstr(sum(n[i,j,k]*items.loc[k,'storage_size'] \
                      for j in J for k in K)<=fc.loc[i,'capacity'])

    mod.setParam('outputflag',False)
    mod.optimize()
    
    writer=pd.ExcelWriter(outputFile)
    pd.DataFrame([mod.objval],columns=['Objective Value'])\
    .to_excel(writer,sheet_name='Summary',index=None)
    
    solution_df = []
    for i in I:
        for j in J:
             for k in K:
                    if not n[i,j,k].x:
                        continue
                    solution_df.append([i, j, k, n[i,j,k].x])
    pd.DataFrame(np.array(solution_df),\
                    columns=['FC_name','region_ID','item_ID','shipment'])\
                    .to_excel(writer,sheet_name='Solution',index=None)
    
    #import heapq
    #heap = []
    capacity_df = []
    for facility in capacity_dict:
        #heapq.heappush(heap, [capacity_dict[facility].pi, facility])
        capacity_df.append([facility, capacity_dict[facility].pi])
    pd.DataFrame(np.array(capacity_df),\
                    columns=['FC_name','shadow_price']).sort_values('shadow_price')\
                    .to_excel(writer,sheet_name='Capacity Constraints',index=None)
    writer.save()
    
    #top = 5
    #top_facilities = []
    #for _ in range(top):
        #_, facility = heapq.heappop(heap)
        #top_facilities.append(facility)

    
    return #top_facilities

if __name__=='__main__':
    import sys, os
    if len(sys.argv)!=3:
        print('Correct syntax: python lab2_code.py inputFile outputFile')
    else:
        inputFile=sys.argv[1]
        outputFile=sys.argv[2]
        if os.path.exists(inputFile):
            optimize(inputFile,outputFile)
            print(f'Successfully optimized. Results in "{outputFile}"')
        else:
            print(f'File "{inputFile}" not found!')
