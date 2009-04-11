#include <stdio.h>
#include <stdlib.h>
#include "sr.h"

/* ******************************************************************
   Selective Repeat NETWORK EMULATOR

   This code should be used for assignment 3, Selective Repeat reliable data transfer protocols (from A to B).  
   Network properties:
   - one way network delay averages five time units (longer if there
   are other messages in the channel)
   - packets can be corrupted (either the header or the data portion)
   or lost, according to user-defined probabilities
   - packets will be delivered in the order in which they were sent
   (although some can be lost).
 **********************************************************************/


/*****************************************************************
 ***************** NETWORK EMULATION CODE STARTS BELOW ***********
 The code below emulates the layer 3 and below network environment:
 - emulates the tranmission and delivery (possibly with bit-level corruption
 and packet loss) of packets across the layer 3/4 interface
 - handles the starting/stopping of a timer, and generates timer
 interrupts (resulting in calling the timer handler implemented by students).
 - generates messages to be sent (passed from later 5 to 4)

 YOU SHOLD NOT TOUCH ANY OF THE CODE BELOW.  
 If you're interested in how the emulator is designed,
 you're welcome to look at the code - but again, you defeinitely 
 should not have to modify
 ******************************************************************/

struct event {
    float evtime;           /* event time */
    int evtype;             /* event type code */
    int eventity;           /* entity where event occurs */
    int seqNum;				/*to record the sequence number, if the event is for time_out*/
    struct pkt *pktptr;     /* ptr to packet (if any) assoc w/ this event */
    struct event *prev;
    struct event *next;
};
struct event *evlist = NULL;   /* the event list */

/* possible events: */
#define  TIMER_INTERRUPT 0  
#define  FROM_LAYER5     1
#define  FROM_LAYER3     2

#define  OFF             0
#define  ON              1
#define   A    0
#define   B    1



int TRACE = 1;             /* for my debugging */
int nsim = 0;              /* number of messages from 5 to 4 so far */ 
int nsimmax = 0;           /* number of msgs to generate, then stop */
float time = 0.000;
float lossprob;            /* probability that a packet is dropped  */
float corruptprob;         /* probability that one bit is packet is flipped */
float lambda;              /* arrival rate of messages from layer 5 */   
int   ntolayer3;           /* number sent into layer 3 */
int   nlost;               /* number lost in media */
int ncorrupt;              /* number corrupted by media*/
FILE * pFile;

void init();
float jimsrand();
void generate_next_arrival();
void insertevent(struct event *p);
void printevlist();

void stoptimer(int AorB, int seqNum);
void starttimer(int AorB,float increment, int seqNum);
void tolayer3(int AorB,struct pkt packet);
void tolayer5(int AorB,  char datasent[20]);
int CheckSum(struct pkt packet);

int main(int argc, char* argv[])
{
    struct event *eventptr;
    struct msg  msg2give;
    struct pkt  pkt2give;

    int i,j;
    char c; 

    init();
    A_init();
    B_init();

    while (1) {
        eventptr = evlist;            /* get next event to simulate */
        if (eventptr==NULL)
            goto terminate;
        evlist = evlist->next;        /* remove this event from event list */
        if (evlist!=NULL)
            evlist->prev=NULL;
        if (TRACE>=2) {
            fprintf(pFile,"\nEVENT time: %f,",eventptr->evtime);
            fprintf(pFile,"  type: %d",eventptr->evtype);
            if (eventptr->evtype==0)
                fprintf(pFile,", timerinterrupt  ");
            else if (eventptr->evtype==1)
                fprintf(pFile,", fromlayer5 ");
            else
                fprintf(pFile,", fromlayer3 ");
            fprintf(pFile," entity: %d\n",eventptr->eventity);
        }
        time = eventptr->evtime;        /* update time to next event time */
        if (nsim==nsimmax)
            break;                        /* all done with simulation */
        if (eventptr->evtype == FROM_LAYER5 ) {
            generate_next_arrival();   /* set up future arrival */
            /* fill in msg to give with string of same letter */    
            j = nsim % 26; 
            for (i=0; i<20; i++)  
                msg2give.data[i] = 97 + j;
            if (TRACE>2) {
                fprintf(pFile,"          MAINLOOP: data given to student: ");
                for (i=0; i<20; i++) 
                    fprintf(pFile,"%c", msg2give.data[i]);
                fprintf(pFile,"\n");
            }
            nsim++;
            if (eventptr->eventity == A) 
                A_rdtsend(msg2give);  
        }
        else if (eventptr->evtype ==  FROM_LAYER3) {
            pkt2give.seqnum = eventptr->pktptr->seqnum;
            pkt2give.acknum = eventptr->pktptr->acknum;
            pkt2give.checksum = eventptr->pktptr->checksum;
            for (i=0; i<20; i++)  
                pkt2give.payload[i] = eventptr->pktptr->payload[i];
            if (eventptr->eventity ==A)      /* deliver packet by calling */
                A_rcv(pkt2give);            /* appropriate entity */
            else
                B_rcv(pkt2give);
            free(eventptr->pktptr);          /* free the memory for packet */
        }
        else if (eventptr->evtype ==  TIMER_INTERRUPT) {
            if (eventptr->eventity == A) 
                A_timerinterrupt(eventptr->seqNum);
        }
        else  {
            fprintf(pFile,"INTERNAL PANIC: unknown event type \n");
        }
        free(eventptr);
    }

terminate:
    fprintf(pFile," Simulator terminated at time %f\n after sending %d msgs from layer5\n",time,nsim);

    exit(1);
}

void init()                         /* initialize the simulator */
{
    int i;
    double sum, avg;
    float jimsrand();
    pFile = fopen("result.txt","w");


    printf("-----  Selective Repeat Network Emulator Version 1.1 -------- \n\n");
    printf("Enter the number of messages to emulate: ");
    scanf("%d",&nsimmax);
    printf("Enter  packet loss probability [enter 0.0 for no loss]:");
    scanf("%f",&lossprob);
    printf("Enter packet corruption probability [0.0 for no corruption]:");
    scanf("%f",&corruptprob);
    printf("Enter average time between messages from sender's layer 5 [ > 0.0]:");
    scanf("%f",&lambda);
    printf("Enter TRACE:");
    scanf("%d",&TRACE);

    srand(9999);              /* init random number generator */
    sum = 0.0;                /* test random number generator for students */
    for (i=0; i<1000; i++)
        sum=sum+jimsrand();    /* jimsrand() should be uniform in [0,1] */
    avg = sum/1000.0;
    if (avg < 0.25 || avg > 0.75) {
        fprintf(pFile,"It is likely that random number generation on your machine\n" ); 
        fprintf(pFile,"is different from what this emulator expects.  Please take\n");
        fprintf(pFile,"a look at the routine jimsrand() in the emulator code. Sorry. \n");
        //exit(1);
    }

    ntolayer3 = 0;
    nlost = 0;
    ncorrupt = 0;

    time=0.0;                    /* initialize time to 0.0 */
    generate_next_arrival();     /* initialize event list */
}

/****************************************************************************/
/* jimsrand(): return a float in range [0,1].  The routine below is used to */
/* isolate all random number generation in one location.					*/
/****************************************************************************/
float jimsrand() 
{
    float x;                   

    x=(float)(rand()%100)/100;		/* x should be uniform in [0,1] */
    return(x);
}  


/********************* EVENT HANDLINE ROUTINES *******/
/*  The next set of routines handle the event list   */
/*****************************************************/

void generate_next_arrival()
{
    double x,log(),ceil();
    struct event *evptr;
    char *malloc();
    float ttime;
    int tempint;

    if (TRACE>2)
        fprintf(pFile,"          GENERATE NEXT ARRIVAL: creating new arrival\n");

    x = lambda*jimsrand()*2;  /* x is uniform on [0,2*lambda] */
    /* having mean of lambda        */
    //evptr = (struct event *)malloc(sizeof(struct event));
    evptr = new struct event;
    evptr->evtime =  time + x;
    evptr->evtype =  FROM_LAYER5;

    evptr->eventity = A;

    insertevent(evptr);
} 


void insertevent(struct event *p)
{
    struct event *q,*qold;

    if (TRACE>2) {
        fprintf(pFile,"            INSERTEVENT: time is %lf\n",time);
        fprintf(pFile,"            INSERTEVENT: future time will be %lf\n",p->evtime); 
    }
    q = evlist;     /* q points to header of list in which p struct inserted */
    if (q==NULL) {   /* list is empty */
        evlist=p;
        p->next=NULL;
        p->prev=NULL;
    }
    else {
        for (qold = q; q !=NULL && p->evtime > q->evtime; q=q->next)
            qold=q; 
        if (q==NULL) {   /* end of list */
            qold->next = p;
            p->prev = qold;
            p->next = NULL;
        }
        else if (q==evlist) { /* front of list */
            p->next=evlist;
            p->prev=NULL;
            p->next->prev=p;
            evlist = p;
        }
        else {     /* middle of list */
            p->next=q;
            p->prev=q->prev;
            q->prev->next=p;
            q->prev=p;
        }
    }
}

void printevlist()
{
    struct event *q;
    int i;
    fprintf(pFile,"--------------\nEvent List Follows:\n");
    for(q = evlist; q!=NULL; q=q->next) {
        fprintf(pFile,"Event time: %f, type: %d entity: %d\n",q->evtime,q->evtype,q->eventity);
    }
    fprintf(pFile,"--------------\n");
}



/**************STUDENTS CAN CALL THE FOLLOWING ROUTINES **********/

/* called by students routine to cancel a previously-started timer */
void stoptimer(int AorB, int seqNum) 
{
    struct event *q,*qold;

    if (TRACE>2)
        fprintf(pFile,"          STOP TIMER: stopping timer at %f\n",time);
    /* for (q=evlist; q!=NULL && q->next!=NULL; q = q->next)  */
    for (q=evlist; q!=NULL ; q = q->next) 
        if ( (q->evtype==TIMER_INTERRUPT  && q->eventity==AorB && q->seqNum == seqNum) ) { 
            /* remove this event */
            if (q->next==NULL && q->prev==NULL)
                evlist=NULL;         /* remove first and only event on list */
            else if (q->next==NULL) /* end of list - there is one in front */
                q->prev->next = NULL;
            else if (q==evlist) { /* front of list - there must be event after */
                q->next->prev=NULL;
                evlist = q->next;
            }
            else {     /* middle of list */
                q->next->prev = q->prev;
                q->prev->next =  q->next;
            }
            free(q);
            return;
        }
    fprintf(pFile,"Warning: unable to cancel your timer. It wasn't running.\n");
}

/* called by students routine to start a timer */
void starttimer(int AorB,float increment,int seqNum) 
{

    struct event *q;
    struct event *evptr;
    char *malloc();

    if (TRACE>2)
        fprintf(pFile,"          START TIMER: starting timer at %f\n",time);
    /* be nice: check to see if timer is already started, if so, then  warn */
    /* for (q=evlist; q!=NULL && q->next!=NULL; q = q->next)  */
    for (q=evlist; q!=NULL ; q = q->next)  
        if ( (q->evtype==TIMER_INTERRUPT  && q->eventity==AorB && q->seqNum == seqNum) ) { 
            fprintf(pFile,"Warning: attempt to start a timer that is already started\n");
            return;
        }

    /* create future event for when timer goes off */
    //evptr = (struct event *)malloc(sizeof(struct event));
    evptr = new struct event;
    evptr->evtime =  time + increment;
    evptr->evtype =  TIMER_INTERRUPT;
    evptr->eventity = AorB;
    evptr->seqNum = seqNum;
    insertevent(evptr);
} 


/*send a packet to layer 3*/
void tolayer3(int AorB,struct pkt packet)
    /* A or B is trying to stop timer */
{
    struct pkt *mypktptr;
    struct event *evptr,*q;
    char *malloc();
    float lastime, x, jimsrand();
    int i;


    ntolayer3++;

    /* simulate losses: */
    if (jimsrand() < lossprob)  {
        nlost++;
        if (TRACE>0)    
            fprintf(pFile,"          TOLAYER3: packet being lost\n");
        return;
    }  

    /* make a copy of the packet student just gave me since he/she may decide */
    /* to do something with the packet after we return back to him/her */ 
    //mypktptr = (struct pkt *)malloc(sizeof(struct pkt));
    mypktptr = new struct pkt;
    mypktptr->seqnum = packet.seqnum;
    mypktptr->acknum = packet.acknum;
    mypktptr->checksum = packet.checksum;
    for (i=0; i<20; i++)
        mypktptr->payload[i] = packet.payload[i];
    if (TRACE>2)  {
        fprintf(pFile,"          TOLAYER3: seq: %d, ack %d, check: %d ", mypktptr->seqnum,
                mypktptr->acknum,  mypktptr->checksum);
        for (i=0; i<20; i++)
            fprintf(pFile,"%c",mypktptr->payload[i]);
        fprintf(pFile,"\n");
    }

    /* create future event for arrival of packet at the other side */
    //evptr = (struct event *)malloc(sizeof(struct event));
    evptr = new struct event;
    evptr->evtype =  FROM_LAYER3;   /* packet will pop out from layer3 */
    evptr->eventity = (AorB+1) % 2; /* event occurs at other entity */
    evptr->pktptr = mypktptr;       /* save ptr to my copy of packet */
    /* finally, compute the arrival time of packet at the other end.
       medium can not reorder, so make sure packet arrives between 1 and 10
       time units after the latest arrival time of packets
       currently in the medium on their way to the destination */
    evptr->seqNum = 0;
    lastime = time;
    /* for (q=evlist; q!=NULL && q->next!=NULL; q = q->next) */
    for (q=evlist; q!=NULL ; q = q->next) 
        if ( (q->evtype==FROM_LAYER3  && q->eventity==evptr->eventity) ) 
            lastime = q->evtime;
    evptr->evtime =  lastime + 1 + 9*jimsrand();



    /* simulate corruption: */
    if (jimsrand() < corruptprob)  {
        ncorrupt++;
        if ( (x = jimsrand()) < .75)
            mypktptr->payload[0]='Z';   /* corrupt payload */
        else if (x < .875)
            mypktptr->seqnum = 999999;
        else
            mypktptr->acknum = 999999;
        if (TRACE>0)    
            fprintf(pFile,"          TOLAYER3: packet being corrupted\n");
    }  

    if (TRACE>2)  
        fprintf(pFile,"          TOLAYER3: scheduling arrival on other side\n");
    insertevent(evptr);
} 


/*deliver data to layer 5*/
void tolayer5(int AorB,  char datasent[20])
{
    int i;  
    if (TRACE>2) {
        fprintf(pFile,"          TOLAYER5: data received: ");
        for (i=0; i<20; i++)  
            fprintf(pFile,"%c",datasent[i]);
        fprintf(pFile,"\n");
    }

}

//calculate the check sum for packet
int CheckSum(struct pkt packet)
{
    struct pkt * addr = &packet;

    long sum = 0;

    sum += packet.seqnum;
    sum += packet.acknum;

    for(int i = 0; i < 20; i++)
        sum += (unsigned short)packet.payload[i];

    while(sum >> 16)
        sum = (sum & 0xffff)+(sum >> 16);

    return ~sum;
}
