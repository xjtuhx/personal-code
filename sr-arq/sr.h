/* ******************************************************************
   Selective Repeat header file

**********************************************************************/

#ifndef	_SR_H
#define	_SR_H


/*window size at both sender and receiver*/
#define WINDOWSIZE 8

/* A "msg" is the data unit passed from layer 5 to layer 4. */
/* It contains the data to be delivered */
struct msg {
  char data[20];
  };

/* A "pkt" is the data unit passed from layer 4 to layer 3.*/
struct pkt {
   int seqnum;
   int acknum;
   int checksum;
   char payload[20];
    };

/* A "buffer"" structure that can be used at both sender and receiver.*/
struct buffer{
	struct pkt packet;
	struct buffer * next;
	struct buffer * pre;
};


/**************STUDENTS CAN CALL THE FOLLOWING ROUTINES **********/

/* called by students routine to cancel a previously-started timer */
void stoptimer(int AorB, int seqNum);

/* called by students routine to start a timer */
void starttimer(int AorB,float increment,int seqNum);

/*send a packet to layer 3*/
void tolayer3(int AorB,struct pkt packet);

/*deliver data to layer 5*/
void tolayer5(int AorB,  char datasent[20]);

//calculate the check sum for packet
int CheckSum(struct pkt packet);

/********* STUDENTS SHOULD WRITE THE NEXT SIX ROUTINES *********/

/* called from layer 5, passed the data to be sent to the other side */
void A_rdtsend(struct msg message);


/* called from layer 3, when a packet arrives for layer 4 */
void A_rcv(struct pkt packet);

/* called when A's timer goes off */
void A_timerinterrupt(int seqNum);

/* the following routine will be called once (only) before any other */
/* A's routines are called. You can use it to do any initialization tasks. */
void A_init();


/* called from layer 3, when a packet arrives for layer 4 at B*/
void B_rcv(struct pkt packet);


/* the following routine will be called once (only) before any other */
/* B's routines are called. You can use it to do any initialization tasks.*/
void B_init();

#endif



