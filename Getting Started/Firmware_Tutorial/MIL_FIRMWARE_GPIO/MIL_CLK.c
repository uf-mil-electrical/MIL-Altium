/*
 * Name: MIL_CLK.h
 * Author: Marquez Jones
 * Date Created: 3/13/19
 * Desc: Clock configuration functions for MIL
 *
 * What to understand: These functions configure your system clock
 *                     in effect, every MCU requires some kind of clock
 *                     source to synchronize the entire sytem
 *
 *                     The TIVA(and all MCUs) provide us with
 *                     multiple possible sources of a clock source
 *                     this can be an external crystal or the result
 *                     of the TIVA's internal clock division circuit
 *
 *                     For the sake of board simplicity ,we will primarily
 *                     depend on the internal oscillator.
 *
 *                     If a design for some reason absolutely needs an
 *                     external oscillator,it will be discussed
 */
#include <stdint.h>
#include <stdbool.h>
#include "driverlib/sysctl.h"

#include "MIL_CLK.h"

/*
 * Name: MIL_ClkSetInt_16MHz
 * Desc: configures the systems clock to
 *       use internal oscillator at 16 MHz
 *
 */
void MIL_ClkSetInt_16MHz(void){

    /*
     * use 16MHz internal oscillator(see page 487 of TivaWare manu)
     * use the oscillator directly( as opposed to the PLL clock div circuit)
     * desired frequency is 16 MHz
     */
    SysCtlClockFreqSet(SYSCTL_OSC_INT |
                       SYSCTL_USE_OSC,
                       MIL_16MHz);

}

