/*
 * Name: MIL_GPIO_Blink
 * Author: Marquez Jones
 * Desc: This is a demo on how to use GPIO
 *       In this demo, we will use one of the push buttons on the launchpad to
 *       toggle the GPIO
 *       These buttons are not debounced in hardware so a live demo may
 *       encounter some unpredictable results due to bouncing
 *       In general hardware designs should have some form of debouncing for push buttons
 *
 *       This demo will turn on the LED if the button is held
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

//defines
#define BLUE_LED_PIN GPIO_PIN_2

//both of these are the same value
//but for sake of clarity I made two separate
//defines
#define PUSH_SW_1_PIN GPIO_PIN_4
#define PUSH_SW_1_bm GPIO_PIN_4

//macros
#define PORTF_CLK_ENABLE() SysCtlPeripheralEnable(SYSCTL_PERIPH_GPIOF);

//function prototypes
void InitSwitch(void);

void InitBlueLED(void);

//Global variabls
uint8_t STATE = 0x00;

/*************************************** MAIN *****************************************/

int main(void)
{

    MIL_ClkSetInt_16MHz();

    //both the LED and switch are attached to port F
    PORTF_CLK_ENABLE();

    //gpio configurations
    InitBlueLED();

    InitSwitch();

    while(1){


        /*
         * Button note:
         * the buttons on the launchpad are active low
         */

        //check if button 1 has been pressed
        //if button is not pressed
        //if so turn LED off
        if(GPIOPinRead(GPIO_PORTF_BASE,PUSH_SW_1_PIN) & PUSH_SW_1_bm){

            STATE = 0x00;

        }
        //otherwise the button was pressed
        else{

            STATE = 0xFF;

        }
        GPIOPinWrite(GPIO_PORTF_BASE, BLUE_LED_PIN, STATE);

    }
}

/***************************************FUNCTION DEFINITIONS***************************/


/******************************************
 * Name: InitSwitch
 * Desc: Initializes GPIO_PIN_4 on port F as an
 *       input pin
 ******************************************/
void InitSwitch(void){

    GPIOPinTypeGPIOInput(GPIO_PORTF_BASE, PUSH_SW_1_PIN);

    /*
     * Fundamental Rule of digital circuits,
     * all signals must be high or low
     * Check launchpad schematic to see why
     * the pull up in software is required
     */

    /*
     * Configure the push input to have an interal pull up resistor
     *
     * Notes:the gpio_strength parameter would set the current output to be
     * 2mA if the pin were configured as an output. In this case it has no effect
     * since the pin is an input pin
     */
    GPIOPadConfigSet(GPIO_PORTF_BASE,
                     PUSH_SW_1_PIN,
                     GPIO_STRENGTH_2MA,
                     GPIO_PIN_TYPE_STD_WPU);

}

void InitBlueLED(void){

    GPIOPinTypeGPIOOutput(GPIO_PORTF_BASE,BLUE_LED_PIN);

}




