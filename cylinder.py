# -*- coding: utf-8 -*-
import pickle
def cylinder(ws2,plcpath,hmipath):
    cyl_length=ws2.cell(row=1,column=12).value
    #append wp number for cylinder
    listwp=[]
    for row in ws2.values:
        if row[0]:
            if str(row[0]).isdigit():
                listwp.append(int(row[0]))
    #for cylinder reset function
    wpMin=min(listwp)
    wpMax=max(listwp)
    #cylinder number of each wp
    qualist={}
    while len(listwp)>0:
        listwp.sort(reverse=True)
        a=listwp.pop()
        b=listwp.count(a)+1
        while listwp.count(a)>0:
            listwp.remove(a)
        qualist[str(a)]=b
    
    #create cyl alarm FB-----------------------------------------------
    fo=open('Cyl_Alarm_FB.pkl','rb')
    tempstr=pickle.load(fo)
    fo.close()
    
    fo = open(plcpath+"\\Cyl_Alarm_FB.scl", "w")
    fo.write(tempstr)
    fo.close()   
    
    #create Cyl_Fault------------------------------------------------------
    
    fo = open(plcpath+"\\Cyl_Fault.scl", "w")
    tempstr='''FUNCTION_BLOCK "Cyl_Fault"
{ S7_Optimized_Access := 'TRUE' }
VERSION : 0.1
   VAR_INPUT 
      System_Time {InstructionName := 'DTL'; LibVersion := '1.0'} : DTL;
      Reset : Bool;
   END_VAR

   VAR 
'''
    fo.write(tempstr)
    
    #instance db block
    for row in ws2.values:
        if row[0]:
            if str(row[0]).isdigit():
                fo.write(r'      "'+str(row[0])+'WorkPosCylinder'+str(row[1])+'" : "Cyl_Alarm_FB";\n')
        
    tempstr='''   END_VAR

   VAR_TEMP 
      i : Int;
      j : Int;
      k : Int;
      a : Array[%d..%d, 1..%d] of Bool;
      b : Array[1..16] of Bool;
   END_VAR

   VAR CONSTANT 
      WpMin : Int := %d;
      WpMax : Int := %d;
      Length : Int := %d;
   END_VAR


BEGIN
''' %(wpMin,wpMax,cyl_length*2,wpMin,wpMax,cyl_length)
    fo.write(tempstr)
    
    #each cylinder
    for row in ws2.values:
        if row[2] and row[2]!='名称':
            #cylinder name
            fo.write('	  //'+row[2]+'\n')
            #cylinder homeinput
            if row[6]:
                fo.write(r'	"Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].inHome := %'+row[6]+';\n')
            else:
                fo.write(r'	"Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].inHome := "Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].outHome;\n')
            #cylinder workinput
            if row[7]:
                fo.write(r'	"Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].inWork :=%'+row[7]+';\n')
            else:
                fo.write(r'	"Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+r'].inWork := "Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].outWork;\n')
            
            tempstr='''	#"%dWorkPosCylinder%d"(System_Time:=#System_Time,
	                     Reset:=#Reset,
	                     inHome:="Cylinder_DB".WorkPos[%d].cylinderInfor[%d].inHome,
	                     inWork:="Cylinder_DB".WorkPos[%d].cylinderInfor[%d].inWork,
	                     Set_Time:=T#5s,
	                     outHome:="Cylinder_DB".WorkPos[%d].cylinderInfor[%d].outHome,
	                     outWork:="Cylinder_DB".WorkPos[%d].cylinderInfor[%d].outWork,
	                     errorHome=>#a[%d,%d],
	                     errorWork=>#a[%d,%d]);\n''' %(row[0],row[1],row[0],row[1],row[0],row[1],row[0],row[1],row[0],row[1],row[0],row[1]*2-1,row[0],row[1]*2)
            fo.write(tempstr)
            #cylinder homeoutput
            if row[4]:
                fo.write('	%'+row[4]+' := "Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].outHome;\n')
            else:
                fo.write(r'	"Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+r'].outHome := "Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].outHome;\n')
            #cylinder workoutput
            if row[5]:
                fo.write('	%'+row[5]+' := "Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].outWork;\n')
            else:
                fo.write(r'	"Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+r'].outWork := "Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].outWork;\n')
            
    tempstr='''	
	FOR #k := #WpMin TO #WpMax DO
	    FOR #i := 0 TO #Length/8-1 DO
	        FOR #j := 1 TO 16 DO
	            #b[#j] := #a[#k,#i * 16 + #j];
	            GATHER(IN := #b,
	                   OUT => "Cylinder_DB".WorkPos[#k].cylinderError[#i + 1]);
	        END_FOR;
	    END_FOR;
	END_FOR;
END_FUNCTION_BLOCK'''
    fo.write(tempstr)
    fo.close()
    
    #Cylinder interlock----------------------------------------------------
    fo = open(plcpath+"\\Cyl_Interlock.scl", "w")
    tempstr='''FUNCTION "Cyl_Interlock" : Void
{ S7_Optimized_Access := 'TRUE' }
VERSION : 0.1

BEGIN
'''
    fo.write(tempstr)
    for row in ws2.values:
        if row[0]:
            if str(row[0]).isdigit():
                fo.write('	//'+row[2]+'\n')
                fo.write('	"Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].enHome := "AlwaysTRUE";\n')
                fo.write('	"Cylinder_DB".WorkPos['+str(row[0])+'].cylinderInfor['+str(row[1])+'].enWork := "AlwaysTRUE";\n')
    fo.write('END_FUNCTION')
    fo.close()
    
    #cylinder hmi--------------------------------------------------------
    fo=open('Cyl_HMI.pkl','rb')
    tempstr=pickle.load(fo)
    fo.close()
    
    fo = open(plcpath+"\\Cyl_HMI.scl", "w")
    fo.write(tempstr %(wpMin,wpMax,int(ws2.cell(1,12).value)))
    fo.close()
    
    #create cylinder reset--------------------------------------------------
    fo=open('Cyl_Reset.pkl','rb')
    tempstr=pickle.load(fo)
    fo.close()
    
    fo = open(plcpath+"\\Cyl_Reset.scl", "w")
    fo.write(tempstr %(wpMin,wpMax))
    fo.close()
    
    #create cylinder DB----------------------------------------------------
    fo = open(plcpath+"\\Cylinder_DB.db", "w")
    tempstr='''DATA_BLOCK "Cylinder_DB"
{ S7_Optimized_Access := 'TRUE' }
VERSION : 0.1
NON_RETAIN
   VAR 
      WorkPos : Array[%d..%d] of Struct
         cylinderQua : Int;
         cylinderInfor : Array[1..%d] of Struct
            enHome : Bool;
            enWork : Bool;
            inHome : Bool;
            inWork : Bool;
            outHome : Bool;
            outWork : Bool;
            Start_Time {InstructionName := 'DTL'; LibVersion := '1.0'} : DTL;
            Running_Time : Time;
            Set_Time : Time;
         END_STRUCT;
         cylinderError : Array[1..%d] of Word;   // 一个word8个气缸
      END_STRUCT;
   END_VAR


BEGIN\n''' %(wpMin,wpMax,int(cyl_length),int(cyl_length/8))
    fo.write(tempstr)
    for key,value in qualist.items():
        fo.write('   WorkPos['+key+'].cylinderQua := '+str(value)+';\n')
    fo.write('END_DATA_BLOCK')
    fo.close()
    
    #create cylinder DB HMI-------------------------------------------------
    fo=open('Cylinder_DB_HMI.pkl','rb')
    tempstr=pickle.load(fo)
    fo.close()
    
    fo = open(plcpath+"\\Cylinder_DB_HMI.db", "w")
    fo.write(tempstr)
    fo.close()
    

    