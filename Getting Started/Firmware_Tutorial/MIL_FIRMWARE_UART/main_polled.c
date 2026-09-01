/*
 * Name: MIL_UART_Lib_Demo
 * Date completed: 3/22/2019
 * Author: Marquez Jones
 * Desc: This will demonstrate using UART on the TIVA C using
 *       a MIL abstraction layer
 *
 *       This will demo will echo received input back to a terminal
 *
 *       This will be accomplished using interrupts.
 *       The polled solution is left in comments in the main loop
 *
 * Hardware Notes:
 * UART 1 on Port B
 * PB0 - UART RX
 * PB1 - UART TX
 *
 * NOTE ABOUT UART INTERRUPTS: In order to clear the interrupt
 *                             you must read the data from the buffer
 */
/* INCLUDES */
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

//MIL includes
#include "MIL_CLK.h"
#include "MIL_UART.h"

/************************FLAGS******************************/

/************************FUNCTION PROTOTYPES******************************/

//These GPIOs will be used to indicate that the program is flashed
void InitGPIO(void);

//turn on all 3 of the launchpad LEDs
void  LedsOn(void);

//turn off all 3 of the launchpad LEDs
void LedsOff(void);

//inaccurate software delay at 16MHz clock
void SoftDelay(float seconds);

/************************MAIN******************************/
int main(void)
{

    /*********************CPU INIT START**********************/
    /*CONFIGURE SYSTEM CLOCK TO INTERNAL 16MHZ*/
    MIL_ClkSetInt_16MHz();

    /******************CPUT INIT END***************************/

    /****************UART INIT START**************************/

    //initialize UART
    MIL_InitUART(UART1_BASE, MIL_DEFAULT_BAUD_115K);

    /****************UART INIT END****************************/

    /****************LED INIT START***************************/
    //initialize GPIO
    InitGPIO();

    //confirm that the program is running by turning on these pins
    LedsOn();

    /*****************LED INIT END***************************/

    /***************DATA START*******************************/

     //C string example
     const char cstring[] = "By Marquez Jones";

    /***************DATA END*********************************/

     MIL_UART_OutCString(UART1_BASE,cstring);

     while(1){

         /************POLLED VERSION OF THE CODE*****************/
        if(UARTCharsAvail(UART1_BASE)){

            //read received data
            uint8_t rx_data = UARTCharGet(UART1_BASE);

             //transmit data
            UARTCharPut(UART1_BASE,rx_data);

        }

    }

	//return 0;
}

/************************FUNCTIONS******************************/

void InitGPIO(void){

    //enable peripheral clock
    SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOF);

    //wait for peripheral to be ready
    while(!SysCtlPeripheralReady(SYSCTL_PERIPH_GPIOF));

    //set port F pins as outputs
    GPIOPinTypeGPIOOutput(GPIO_PORTF_BASE, GPIO_PIN_1 | GPIO_PIN_2 |GPIO_PIN_3);


}

void LedsOn(void){

    //confirm that the program is running by turning on these pins
    GPIOPinWrite(GPIO_PORTF_BASE,
                 GPIO_PIN_1 | GPIO_PIN_2 |GPIO_PIN_3,
                 GPIO_PIN_1 | GPIO_PIN_2 |GPIO_PIN_3);

}

void LedsOff(void){

    GPIOPinWrite(GPIO_PORTF_BASE,
                 GPIO_PIN_1 | GPIO_PIN_2 | GPIO_PIN_3,
                 0x00);


}

void SoftDelay(float seconds){

    //at 16MHz clock
    uint32_t cycles = 1600000 * seconds;

    for(uint32_t i = 0; i < cycles; i++){}

}



