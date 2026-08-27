/*
 * Name: MIL_UART
 * Author: Marquez Jones
 * Date created: 3/19/2019
 * Desc: A set of abstraction functions
 *       to allow rapid deployment of UART
 *
 * Note: These functions allow for the most
 *       bare bone application of UART
 *       for more advanced functionality,
 *       you may need to search through the manuals
 *
 * Baud Rate Note:
 *      The default baud rate will be 115.2k
 *      this is what all MIL devices
 *
 *      other baud rates are allowed,but
 *      I don't recommend using them unless
 *      you have a reason to not use the
 *      MIL_DEFAULT
 *
 * Hardware Notes:
 *       UART1 can technically also use PC4/PC5 for RX/TX
 *       which is also shared by UART4 so I did not include
 *       the option to use UART1 on that set of pins
 *
 *       This should not cause any hardware conflicts
 *
 *       All modules function exactly the
 *       except for UART1 which also has
 *       flow of control features
 *
 *       A separate function will be created
 *       to handle cases with this as UART1
 *       is the only module with this feature
 *
 * MIL_UART PIN MAP:
 *      UART0:
 *          RX :  PA0
 *          TX :  PA1
 *      UART1:
 *          RX :  PB0
 *          TX :  PB1
 *      UART2:
 *          RX :  PD6
 *          TX :  PD7
 *      UART3:
 *          RX :  PC6
 *          TX :  PC7
 *      UART4:
 *          RX :  PC4
 *          TX :  PC5
 *      UART5:
 *          RX :  PE4
 *          TX :  PE5
 *      UART6:
 *          RX :  PD4
 *          TX :  PD5
 *      UART7:
 *          RX :  PE0
 *          TX :  PE1
 *
 *      NOTE: AS STATED BEFORE
 *            WE'RE MISSING ONE POSSIBLE
 *            UART PIN SET, UART1 CAN ALSO USE PC4/PC5 AS RX/TX
 *            WHICH IS ALSO USED BY UART4
 */

#include "driverlib/uart.h"
#include "utils/uartstdio.h"
#include <stdint.h>

#ifndef MIL_UART_H_
#define MIL_UART_H_




//useful baud rates including the default
//just a list of standard baud
/*
 * Yes, this is more typing than
 * just writing the number
 * but it specifies that the value
 * is a baud rate as opposed to a magic
 * number
 */
#define MIL_DEFAULT_BAUD_115K 115200    //most applications should use this
#define MIL_BAUD_9600  9600
#define MIL_BAUD_57600 57600
#define MIL_BAUD_38400 38400
#define MIL_BAUD_19200 19200
#define MIL_BAUD_YEET     69420
#define MIL_BAUD_SCHWARTZ 37000

//Ascii defines
//CR and LR get sent when you hit enter on a keyboard
#define CR 0x0D //carriage return
#define LF 0x0A //line feed
#define BS 0x08 //backspace

//Interrupt flag defines
//completely arbitrary
#define MIL_RX_INT_EN UART_INT_RX
#define MIL_TX_INT_EN UART_INT_TX

/*
 * Desc: Enables a specified UART base
 *       at a specified baud rate
 *
 *       configured for:
 *          8 bit words
 *          no parity
 *          one stop bit
 *
 *       Interrupts not inherently enabled
 *       FIFO disabled(must be enabled separately)
 *
 * FIFO Note: With fifo disabled RX and TX interrupts
 *            will occur after one byte. If fifo disabled
 *            interrupts would occur what whatever depth
 *            the FIFO is
 *
 * Parameters:
 *            base: UART TIVA base UARTx_BASE(where x is 0 to 7)
 *            baud_rate: your communication speed(see MIL_BAUD defines)
 *
 * BAUD RATE NOTE:
 *            I recommend using the defines in this file, but you could
 *            technically use any number just as long as everything else
 *            matched that baud rate
 *
 *            The standard baud rate for MIL should be 115.2k unless needed
 *            otherwise
 */
void MIL_InitUART(uint32_t base,uint32_t baud_rate);

/*
 * Desc: This function will enable specified interrupts
 *       for the specified module
 *
 * Parameters:
 *        base: UART TIVA base UARTx_BASE(where x is 0 to 7)
 *        int_flags: see MIL_INT_EN defines for
 *                   this is the bit field for UART interrupts
 *        pISR: pointer to your user defined interrupt service routine
 *
 *        INT FLAGS NOTE: if you set more than one interrupt flag,
 *                        you will have to decode the source in the ISR
 *                        as you cannot give each flag it's own ISR
 *
 *        ISR RECOMMENDATIONS: Most applications in MIL will only require
 *                             an RX interrupt. So this is the only one I
 *                             recommend setting ,but everything is based on
 *                             the needs of the project
 */
void MIL_UART_InitISR(uint32_t base,uint32_t int_flags,void (*pISR)(void));

/*
 * Desc: FIFO is disabled by default in MIL_InitUART
 *       this function enables it and allows you to
 *       determine at what fullness will the trigger
 *       RX/TX interrupts
 *
 * Parameters:
 *       base: your TIVA UART base
 *       int_depth: how many bytes to trigger the interrupt
 *                  THE TIVA ONLY SUPPORTS DEPTHS 1,2,4,6,AND 7
 *                  ANYTHING ELSE IS INVALID
 *
 * NOTE: THE FIFO HAS A DEPTH OF 8, THE DEPTH VARIABLE FOR THIS
 *       FUNCTION DETERMINES WHEN INTERRUPTS GET TRIGGERED
 *
 *       NOT HOW MANY BYTES CAN BE STORED TO THE FIFO
 *
 */
void MIL_UART_FIFOEn(uint32_t base, uint8_t int_depth);

/*
 * Desc: send out a an array of data a predefined length
 *
 * Parameters:
 * base : Tiva UARTx_BASE
 * pMsg : a pointer to your data(note arrays in C are pointers)
 * len  : how many bytes of data are you sending
 *
 */
void MIL_UART_OutArray(uint32_t base, uint8_t *pMsg,uint8_t len);

/*
 * Desc: send out a C string,function will end when the value 0x00 is
 *       detected
 *
 * Note: search the syntax for C strings
 *       basically an array of characters terminated by a null
 *       character is considered a C string
 *
 *       C doesn't have string objects like C++
 *
 * Parameters:
 * base : Tiva UARTx_BASE
 * pMsg : a pointer to your data(note arrays in C are pointers)
 */
void MIL_UART_OutCString(uint32_t base, uint8_t *pMsg);


#endif /* MIL_UART_H_ */
