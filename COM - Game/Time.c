

void U4V(){
//Timer_4++;
//U4++;
//if (U4==3){U4=0;}
//if (Timer_4==0){}
//if (Timer_4==1){U4++;Timer_4=0;}
//if (U4==3){U4=0;}

}


void Time1(){
Timer_3++;
//if (Timer_3==0){}
if (Timer_3==5000){sec++;Timer_3=0;} //x1,y4
if (sec==59){sec=0;min++;}
if (min==59){min=0;hour++;}
if (hour==23){hour=0;}
if (PWM_Susp==1){
//Timer_4++;
//if (Timer_4==100000){PWM_Susp=0;Timer_4=0;}
}
U4V();
//if (PWM_Susp==0){Dig();}
Dig();
}