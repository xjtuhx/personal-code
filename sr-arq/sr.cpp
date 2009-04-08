#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "sr.h"

static struct buffer *A_buffer_head = NULL;
static struct buffer *A_buffer_tail = NULL;
static struct buffer *A_window_head = NULL;
static struct buffer *A_window_tail = NULL;
static struct buffer *B_buffer_head = NULL;
static struct buffer *B_buffer_tail = NULL;
static struct buffer *B_window_head = NULL;
static struct buffer *B_window_tail = NULL;

const unsigned int MAX_SEQ = WINDOWSIZE * 2;

unsigned int A_pkt_sent = 0;

/* called from layer 5, passed the data to be sent to the other side */
void A_rdtsend(struct msg message)
{
	struct buffer *p = (struct buffer)malloc(sizeof(struct buffer));
    static unsigned int seq = 0;
    p->next = NULL;
    p->pre = NULL;
    memcpy(p->packet.payload, message.data, 20);
    p->packet.seqnum = seq;
    seq = (seq + 1) % MAX_SEQ;
    p->packet.acknum = 0;
    p->packet.checksum = CheckSum(p->packet);
    if (A_buffer_head == NULL)
    {
        A_buffer_head = p;
        A_buffer_tail = p;
    } else {
        A_buffer_tail->next = p;
        p->pre = A_buffer_tail;
        A_buffer_tail = p;
    }

    p = A_buffer_head;
    while (pkt_sent < WINDOWSIZE && p != NULL)
    {
        /* send pkt */
        starttimer(0, 20, p->packet.seqnum);
        tolayer3(0, p->packet);
        pkt_sent++;
        p = p->next;
    }
}


/* called from layer 3, when a packet arrives for layer 4 */
void A_rcv(struct pkt packet)
{
	struct buffer * p = NULL;
    int i = 0;
    if (CheckSum(packet) != packet.checksum)
        return;
    for(p = A_buffer_head, i = 0; i < WINDOWSIZE && p != NULL;
            i++, p = p->next)
    {
        if (p == A_buffer_head)
        {
            if (p->packet.seqnum + 1 == packet.acknum
                    || p->packet.seqnum == -1)
            {
                A_buffer_head = p->next;
                free(p);
                p = A_buffer_head;
                pkt_sent--;

}

/* called when A's timer goes off */
void A_timerinterrupt(int seqNum)
{
	
}

/* the following routine will be called once (only) before any other */
/* A's routines are called. You can use it to do any initialization tasks. */
void A_init()
{
	
}


/* called from layer 3, when a packet arrives for layer 4 at B*/
void B_rcv(struct pkt packet)
{
	
}


/* the following routine will be called once (only) before any other */
/* B's routines are called. You can use it to do any initialization tasks.*/
void B_init()
{
	
}

