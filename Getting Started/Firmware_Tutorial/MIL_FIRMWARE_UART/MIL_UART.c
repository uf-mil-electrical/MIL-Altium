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
#include <stdbool.h>
#include <stdint.h>
#include "inc/hw_can.h"
#include "inc/hw_ints.h"
#include "inc/hw_memmap.h"
#include "inc/hw_types.h"
#include "driverlib/can.h"
#include "driverlib/gpio.h"
#include "driverlib/interrupt.h"
#include "driverlib/pin_map.h"
#include "driverlib/sysctl.h"
#include "driverlib/uart.h"
#include "utils/uartstdio.h"

#include"MIL_UART.h"

/*
 * Desc: Enables a specified UART base
 *       at a specified baud rate
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
void MIL_InitUART(uint32_t base,uint32_t baud_rate){

    switch(base){

        //RX :  PA0
        //TX :  PA1
        case UART0_BASE:
            SysCtlPeripheralEnable(SYSCTL_PERIPH_UART0);
            SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOA);
            GPIOPinConfigure(GPIO_PA0_U0RX);
            GPIOPinConfigure(GPIO_PA1_U0TX);
            GPIOPinTypeUART(GPIO_PORTA_BASE, GPIO_PIN_0 | GPIO_PIN_1);
            break;

        //RX :  PB0
        //TX :  PB1
        case UART1_BASE:
            SysCtlPeripheralEnable(SYSCTL_PERIPH_UART1);
            SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOB);
            GPIOPinConfigure(GPIO_PB0_U1RX);
            GPIOPinConfigure(GPIO_PB1_U1TX);
            GPIOPinTypeUART(GPIO_PORTB_BASE, GPIO_PIN_0 | GPIO_PIN_1);
            break;

        //RX :  PD6
        //TX :  PD7
        case UART2_BASE:
            SysCtlPeripheralEnable(SYSCTL_PERIPH_UART2);
            SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOD);
            GPIOPinConfigure(GPIO_PD6_U2RX);
            GPIOPinConfigure(GPIO_PD7_U2TX);
            GPIOPinTypeUART(GPIO_PORTD_BASE, GPIO_PIN_6 | GPIO_PIN_7);
            break;

        //RX :  PC6
        //TX :  PC7
        case UART3_BASE:
            SysCtlPeripheralEnable(SYSCTL_PERIPH_UART3);
            SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOC);
            GPIOPinConfigure(GPIO_PC6_U3RX);
            GPIOPinConfigure(GPIO_PC7_U3TX);
            GPIOPinTypeUART(GPIO_PORTC_BASE, GPIO_PIN_6 | GPIO_PIN_7);
            break;

        //RX :  PC4
        //TX :  PC5
        case UART4_BASE:
            SysCtlPeripheralEnable(SYSCTL_PERIPH_UART4);
            SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOC);
            GPIOPinConfigure(GPIO_PC4_U4RX);
            GPIOPinConfigure(GPIO_PC5_U4TX);
            GPIOPinTypeUART(GPIO_PORTC_BASE, GPIO_PIN_4 | GPIO_PIN_5);
            break;

        //RX :  PE4
        //TX :  PE5
        case UART5_BASE:
            SysCtlPeripheralEnable(SYSCTL_PERIPH_UART5);
            SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOE);
            GPIOPinConfigure(GPIO_PE4_U5RX);
            GPIOPinConfigure(GPIO_PE5_U5TX);
            GPIOPinTypeUART(GPIO_PORTE_BASE, GPIO_PIN_4 | GPIO_PIN_5);
            break;

        //RX :  PD4
        //TX :  PD5
        case UART6_BASE:
            SysCtlPeripheralEnable(SYSCTL_PERIPH_UART6);
            SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOD);
            GPIOPinConfigure(GPIO_PD4_U6RX);
            GPIOPinConfigure(GPIO_PD5_U6TX);
            GPIOPinTypeUART(GPIO_PORTD_BASE, GPIO_PIN_4 | GPIO_PIN_5);
            break;

        //RX :  PE0
        //TX :  PE1
        case UART7_BASE:
            SysCtlPeripheralEnable(SYSCTL_PERIPH_UART7);
            SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOE);
            GPIOPinConfigure(GPIO_PE0_U7RX);
            GPIOPinConfigure(GPIO_PE1_U7TX);
            GPIOPinTypeUART(GPIO_PORTE_BASE, GPIO_PIN_0 | GPIO_PIN_1);
            break;

        default:
            break;

    };

    UARTConfigSetExpClk(base ,
                        SysCtlClockGet(),
                        baud_rate,
                       (UART_CONFIG_WLEN_8 |
                        UART_CONFIG_STOP_ONE |
                        UART_CONFIG_PAR_NONE));

    UARTEnable(base);

    UARTFIFODisable(base);



}

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
void MIL_UART_InitISR(uint32_t base,uint32_t int_flags,void (*pISR)(void)){

    /*INTERRUPTS*/
    //Peripheral Interrupt configs
    UARTIntEnable(base,int_flags);
    UARTIntRegister(base, pISR);

}

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
void MIL_UART_FIFOEn(uint32_t base, uint8_t int_depth){

    uint32_t rxlevel;
    uint32_t txlevel;
    if(int_depth >= 7){

        rxlevel = UART_FIFO_RX7_8;
        txlevel = UART_FIFO_TX7_8;

    }
    else if(int_depth >= 5){

        rxlevel = UART_FIFO_RX6_8;
        txlevel = UART_FIFO_TX6_8;

    }
    else if(int_depth >= 3){

        rxlevel = UART_FIFO_RX4_8;
        txlevel = UART_FIFO_TX4_8;

    }
    else if(int_depth == 2){

        rxlevel = UART_FIFO_RX2_8;
        txlevel = UART_FIFO_TX2_8;

    }
    else{

        rxlevel = UART_FIFO_RX1_8;
        txlevel = UART_FIFO_TX1_8;

    }

    UARTFIFOLevelSet(base,txlevel,rxlevel);

    UARTFIFOEnable(base);
}

/*
 * Desc: send out a an array of data a predefined length
 *
 * Parameters:
 * base : Tiva UARTx_BASE
 * pMsg : a pointer to your data(note arrays in C are pointers)
 * len  : how many bytes of data are you sending
 *
 */
void MIL_UART_OutArray(uint32_t base, uint8_t *pMsg,uint8_t len){

    for(uint8_t i = 0; i < len;i++){ UARTCharPut(base,pMsg[i]);}

}

/*
 * Desc: send out a C string,function will end when the value 0x00 is
 *       detected
 *
 *       WILL TRANSMIT 0x00 AT END
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
void MIL_UART_OutCString(uint32_t base, uint8_t *pMsg){

    do{ UARTCharPut(base,*pMsg); pMsg++;}while(*pMsg);

}


