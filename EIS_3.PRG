USE EIS
APPE BLANK



@  2,  0  SAY ""
@  2, 12  GET  EIS->HME
@  3,  0  SAY "./"
@  3, 12  GET  EIS->APA
@  4,  0  SAY ""
@  4, 12  GET  EIS->APOST
@  5,  0  SAY ""
@  5, 12  GET  EIS->PARAL
@  6,  0  SAY " ."
@  6, 12  GET  EIS->POLH
@  7,  0  SAY "."
@  7, 12  GET  EIS->POSO
@  8,  0  SAY ""
@  8, 12  GET  EIS->FPA
@  9,  0  SAY ""
@  9, 12  GET  EIS->SYNOLO
@ 10,  0  SAY ""
@ 10, 12  GET  EIS->PROMHU
@ 11,  0  SAY "MAKINE"
@ 11, 12  GET  EIS->MAKRINES


READ



PROCEDURE EISAGVGH     && PBNAME2.ONCLICK
*-----------------------------------------------------------------
* NAME
*   ONCLICK - Event handler for PBNAME2
*
*-----------------------------------------------------------------
   REPLACE APOST WITH '///////'
	
RETURN    


PROCEDURE DIORUVSH     && PBNAME3.ONCLICK
*-----------------------------------------------------------------
* NAME
*   ONCLICK - Event handler for PBNAME3
*
*-----------------------------------------------------------------
    R=1
	@ 10,0 SAY '  ' GET R
	READ
RETURN    


PROCEDURE GOTFOCUS_EIS     && APA003.ONGOTFOCUS
*-----------------------------------------------------------------
* NAME
*   ONGOTFOCUS - Event handler for APA003
*
*-----------------------------------------------------------------
	APPEND BLANK
	EDIT
RETURN    


PROCEDURE EIS2     && PBNAME2.ONGOTFOCUS
*-----------------------------------------------------------------
* NAME
*   ONGOTFOCUS - Event handler for PBNAME2
*
*-----------------------------------------------------------------
	APPEND BLANK
	EDIT
	
RETURN    
