/*
 * Name: MIL_GPIO_Blink
 * Author: Marquez Jones
 * Desc: This is an example on how to toggle an
 *       LED using GPIO
 *
 *       This code will:
 *       Initalize GPIO
 *       Toggle it on and off forever
 *
 */

//includes
#include <stdbool.h>
#include <stdint.h>
#include "inc/hw_ints.h"
#include "inc/hw_memmap.h"
#include "driverlib/gpio.h"
#include "driverlib/interrupt.h"
#include "driverlib/pin_map.h"
#include "driverlib/sysctl.h"
#include "driverlib/timer.h"
#include "driverlib/uart.h"
#include "utils/uartstdio.h"

//mil includes
#include "MIL_CLK.h"

/*************************************** DEFINES/MACROS ******************************/

//defines
#define BLUE_LED_PIN GPIO_PIN_2    //check the TM4C123 Launchpad schematic

//macros
#define PORTF_CLK_ENABLE() SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOF);


/*************************************** FUNCTION PROTOTYPES *************************/

/*******************************************
 * Name: InitBlueLED
 * Desc: Set GPIO_PIN_2(BLUE LED) as output
 *******************************************/
void InitBlueLED(void);

/*************************************** MAIN *****************************************/

int main(void)
{
    // initialize clock
    MIL_ClkSetInt_16MHz();

    //both the LED and switch are attached to port F
    PORTF_CLK_ENABLE();

    //gpio configurations
    InitBlueLED();

    while(1){

        //turn LED on

        //software delay(waste time)
        SysCtlDelay(1E6);

        //turn LED off

        //software delay(waste time)
        SysCtlDelay(1E6);

    }
}

/***************************************FUNCTION DEFINITIONS***************************/

/*******************************************
 * Name: InitBlueLED
 * Desc: Set GPIO_PIN_2(BLUE LED) as output
 *******************************************/
void InitBlueLED(void){


}




