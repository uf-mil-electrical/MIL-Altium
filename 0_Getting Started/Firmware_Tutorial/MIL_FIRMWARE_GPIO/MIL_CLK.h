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
 *
 * Clock system diagram for TIVA:
 * see page 222 ,figure 5-5 of Tiva MCU manual to see how
 * clock system in connected
 */

#ifndef MIL_CLK_H_
#define MIL_CLK_H_

#define MIL_16MHz 16000000

/*
 * Name: MIL_ClkSetInt_16MHz
 * Desc: configures the systems clock to
 *       use internal oscillator at 16 MHz
 */
void MIL_ClkSetInt_16MHz(void);


#endif /* MIL_CLK_H_ */
