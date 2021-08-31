# -*- coding: utf-8 -*-

import pickle

temp='''FUNCTION_BLOCK "Cyl_Alarm_FB"
{ S7_Optimized_Access := 'TRUE' }
VERSION : 0.1
   VAR_INPUT 
      System_Time {InstructionName := 'DTL'; LibVersion := '1.0'; ExternalAccessible := 'False'; ExternalVisible := 'False'; ExternalWritable := 'False'} : DTL;
      Reset { ExternalAccessible := 'False'; ExternalVisible := 'False'; ExternalWritable := 'False'} : Bool;
      inHome { ExternalAccessible := 'False'; ExternalVisible := 'False'; ExternalWritable := 'False'} : Bool;
      inWork { ExternalAccessible := 'False'; ExternalVisible := 'False'; ExternalWritable := 'False'} : Bool;
      Set_Time : Time;
   END_VAR

   VAR_OUTPUT 
      errorHome : Bool;
      errorWork : Bool;
   END_VAR

   VAR_IN_OUT 
      outHome { ExternalAccessible := 'False'; ExternalVisible := 'False'; ExternalWritable := 'False'} : Bool;
      outWork { ExternalAccessible := 'False'; ExternalVisible := 'False'; ExternalWritable := 'False'} : Bool;
   END_VAR

   VAR 
      Start_Time {InstructionName := 'DTL'; LibVersion := '1.0'} : DTL;
      Running_Time : Time;
   END_VAR


BEGIN
	IF ((#inHome XOR #outHome) OR (#inWork XOR #outWork)) AND (#outHome OR #outWork) THEN
	    #Running_Time := T_DIFF(IN1 := #System_Time, IN2 := #Start_Time);
	    IF #Running_Time >= #Set_Time THEN
	        
	        #errorHome := (#inHome XOR #outHome) AND #outHome;
	        
	        #errorWork := (#inWork XOR #outWork) AND #outWork;
	        
	        #outWork := FALSE;
	        #outHome := FALSE;
	    END_IF;
	ELSE
	    #Start_Time := #System_Time;
	END_IF;
	
	IF #Reset THEN
	    #errorHome := FALSE;
	    #errorWork := FALSE;
	END_IF;
	
END_FUNCTION_BLOCK'''

fo=open('Cyl_Alarm_FB.pkl','wb')
pickle.dump(temp,fo)
fo.close()

'''
fo=open('Cylinder_DB_HMI.pkl','rb')
temp=pickle.load(fo)
fo.close()

fo=open('cylindertest.db','w')
fo.write(temp)
fo.close()

b=a=1
row=[1,1,'aaa','bbb','','Q1.0','','','T']

print([str(a),'Discrete_alarm_'+str(a),row[2]+'缩回报警'+row[6],'','Warnings','Cylinder_Error',str(b+8),'<No value>','0','<No value>','0','<No value>','False','<No value>'])
'''