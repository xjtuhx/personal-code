#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "sr.h"

static struct buffer *A_buffer_head = NULL;
static struct buffer *A_buffer_tail = NULL;
static struct buffer *A_window_tail = NULL;
static struct buffer *B_buffer_head = NULL;

unsigned int A_pkt_sent = 0;

/* called from layer 5, passed the data to be sent to the other side */
void A_rdtsend(struct msg message)
{
    struct buffer *p = (struct buffer*)malloc(sizeof(struct buffer));
    static int seq = 0;
    p->next = NULL;
    p->pre = NULL;
    memcpy(p->packet.payload, message.data, 20);
    p->packet.seqnum = seq;
    seq++;
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

    /* Send packet */
    if (A_pkt_sent < WINDOWSIZE) 
    {
        starttimer(0, 20, p->packet.seqnum);
        tolayer3(0, p->packet);
        A_pkt_sent++;
        A_window_tail = p;
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
                    || p->packet.acknum == -1)
            {
                if (p->packet.acknum != -1)
                    stoptimer(0, p->packet.seqnum);
                A_buffer_head = p->next;
                free(p);
                p = A_buffer_head;
                if (A_window_tail != NULL && A_window_tail->next != NULL)
                {
                    A_window_tail = A_window_tail->next;
                    starttimer(0, 20, A_window_tail->packet.seqnum);
                    tolayer3(0, A_window_tail->packet);
                } else {
                    A_pkt_sent--;
                }
            }
        } else if (p->packet.seqnum + 1 == packet.acknum) {
            if (p->packet.acknum != -1)
                stoptimer(0, p->packet.seqnum);
            p->packet.acknum = -1;
        }
    }
}

/* called when A's timer goes off */
void A_timerinterrupt(int seqNum)
{
    struct buffer * p = NULL;
    int i = 0;
    for (p = A_buffer_head, i = 0; i < WINDOWSIZE && p != NULL;
            i++, p = p->next)
    {
        if (p->packet.seqnum == seqNum
                && p->packet.acknum != -1)
        {
            starttimer(0, 20, seqNum);
            tolayer3(0, p->packet);
            break;
        }
    }
}

/* the following routine will be called once (only) before any other */
/* A's routines are called. You can use it to do any initialization tasks. */
void A_init()
{

}


/* called from layer 3, when a packet arrives for layer 4 at B*/
void B_rcv(struct pkt packet)
{
    struct buffer * p = NULL;
    struct buffer * head = NULL;
    int i = 0;
    static int nRecvd = 0;
    if (CheckSum(packet) != packet.checksum)
        return;
    p = (struct buffer*)malloc(sizeof(struct buffer));
    memcpy(p->packet.payload, packet.payload, 20);
    p->packet.seqnum = packet.seqnum;
    p->packet.acknum = packet.acknum;
    p->packet.checksum = packet.checksum;
    p->next = NULL;
    p->pre = NULL;


    for (head = B_buffer_head, i = 0;
            i < WINDOWSIZE && head != NULL;
            i++, head = head->next)
    {
        if (head->packet.seqnum > p->packet.seqnum)
        {
            p->next = head;
            if (head->pre != NULL)
                head->pre->next = p;
            p->pre = head->pre;
            head->pre = p;
            nRecvd++;
            packet.acknum = packet.seqnum + 1;
            packet.checksum = CheckSum(packet);
            tolayer3(1, packet);
            break;
        } else if (head->next == NULL) {
            /* Queue tail */
            head->next = p;
            p->pre = head;
            packet.acknum = packet.seqnum + 1;
            packet.checksum = CheckSum(packet);
            tolayer3(1, packet);
            nRecvd++;
            break;
        }
    }

    if (B_buffer_head == NULL)
    {
        B_buffer_head = p;
        nRecvd++;
        packet.acknum = packet.seqnum + 1;
        packet.checksum = CheckSum(packet);
        tolayer3(1, packet);
    }

    if (nRecvd == WINDOWSIZE)
    {
        for (head = B_buffer_head, i = 0;
                i < WINDOWSIZE && head != NULL;
                i++)
        {
            tolayer5(1, head->packet.payload);
            p = head->next;
            free(head);
            head = p;
        }
    }
}


/* the following routine will be called once (only) before any other */
/* B's routines are called. You can use it to do any initialization tasks.*/
void B_init()
{

}

