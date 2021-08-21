


SET CONFIRM ON
set dele on
set date british

publ ektypoths,updating
updating=1
ektypoths=1
buff = 'e'
sele 1
use eis
if .not. file ( 'APA.NTX' )
   INDEX ON eidos+apa TO APA
ENDIF
if .not. file ( 'HME.NTX' )
   INDEX ON eidos+dtos(hme) TO HME
ENDIF
SET COLOR TO W/B,B/W
sele 2
  use \lageuro\pel index \lageuro\ko,\lageuro\ep
sele 1

BASTIM_GID = 2
DO CASE
   CASE BASTIM_GID = 1  && COLOR CARD
      WN='RB/N'
      NW='N/BG'
      WPN='BG+/N'
      wnbl='RB*/N'
   CASE BASTIM_GID = 0
      WN='W/N'
      NW='N/W'
      WPN='W+/N'
      wnbl='W*/N'
   CASE BASTIM_GID = 2
      WN='W/B'
      NW='R/W'
      WPN='BG+/N'
      wnbl='W*/B'
   CASE BASTIM_GID = 3
      WN='w/r'
      NW='r/w'
      WPN='w+/r'
      wnbl='w*/r'
ENDCASE


CLEAR
EPILOGH=1

DO WHILE .T.

     sele 1
        use eis index apa,hme
     sele 2
        use \lageuro\pel index \lageuro\ko,\lageuro\ep
     sele 1






 CLEAR
 DO EPILOGH3 WITH 10,10,EPILOGH,6,'ÑàëÑêïéãÑåÄ','ÑçÑêïéãÑåÄ','ANAZHTHëÜ îéêíóíàâÜë','èÄêéìëàÄëÜ','ëíéàï.èÄêÄÉéëÜë','ÄåÄâÄíÄëâÑìÜ INDEX'

  DO CASE
     CASE EPILOGH=1
         DO EISERX
     CASE EPILOGH=2
         DO EJERX
     CASE EPILOGH=3
         DO ANAZHT
     CASE EPILOGH=4
         DO PAROYS
     CASE EPILOGH=5
         do stoix_parad
     CASE EPILOGH=6
         ! DEL *.NTX
         if .not. file ( 'APA.NTX' )
            INDEX ON eidos+apa TO APA
         ENDIF
         if .not. file ( 'HME.NTX' )
            INDEX ON eidos+dtos(hme) TO HME
         ENDIF
     OTHERWISE
         QUIT
  ENDCASE


ENDDO

RETURN



PROC PAROYS
***************************
CLEAR
STORE DATE() TO HM1,HM2
TYPOS=3
M_apos ='          '
M_paral='          '
M_etairia ='  '
m_eidos='ë'
m_antikat='*'
m_prom='*'
@ 1 ,10 say 'ÑàëÑêïéãÑåÄ=ë ÑçÑêïéãÑåÄ=ç  âÄà íÄ Éìé=*  ' get m_eidos  valid m_eidos $ 'ëç*'
@ 2,10 SAY 'Äèé ' get hm1
@ 4,10 SAY 'Äèé ' get hm2
@ 6,10 say '1=collect 2=ß†©´‡©û 3=¶¢ò' get typos pict '9'
@ 7,10 say '1=ãéåé ÄåíàâÄíÄÅéäÑë  *=éäÄ ' get m_ANTIKAT   pict '9'
@ 8,10 say '1=£ú ß®¶£„üú†ò        *=éäÄ ' get m_prom   pict '9'
read

if  m_eidos = 'ë'
    @ 10,10 say 'èÄêÄäÜèíÜë ' get m_paral   valid intpar('pel','kod','epo','m_paral')=0
endif

if m_eidos = 'ç'
    @ 8,10 say  'ÄèéëíéäÑÄë ' get m_apos    valid intpar('pel','kod','epo','m_apos')=0
endif

if m_eidos = '*'
    @ 8,10 say  'ÄèéëíéäÑÄë „ èÄêÄäÜèíÜë ' get m_apos    valid intpar('pel','kod','epo','m_apos')=0
endif

m_taj='*'
KOLL=1


@ 11,10 say '´ò•†§Ê£û©û *=û£ú®¶£û§Âò 1=ßò®ò¢ 2=òß¶©´  ' get m_taj

@ 12,10 say 'ÑíÄàêàÄ    ' get m_etairia

@ 13 ,10 say 'ëíÜäÑë ãÑ COLLECT / PPD =1 AääÑë = 0 ' get KOLL
read
SYNT='HME>=HM1 .AND. HME<=HM2'
IF TYPOS=1
   SYNT=SYNT+'.AND. COLLECT $ "Nn§å" '
ENDIF
IF TYPOS=2
   SYNT=SYNT+'.AND. PISTOSH $ "Nn§å" '
ENDIF
IF M_apos <> '  '
   if m_eidos='*'
      SYNT=SYNT+'.AND. ( trim(m_apos)=TRIM(APOST) .or. trim(m_apos)=TRIM(PARAL) )'
   else
      SYNT=SYNT+'.AND. trim(m_apos)=TRIM(APOST) '
   endif
ENDIF
IF M_paral <> '  '
   SYNT=SYNT+'.AND. trim(m_paral)=TRIM(PARAL) '
ENDIF
IF M_etairia <> '  '
   SYNT=SYNT+'.AND. etairia = trim(m_etairia) '
ENDIF

if m_ANTIKAT='1'
   SYNT=SYNT+'.AND. antikat>0 '
endif

if m_prom='1'
   SYNT=SYNT+'.AND. promhu>0 '
endif

 p=1
    do epilogh3 with 12,10,p,2,'ú°´Áß‡©û ©ú ¶üÊ§û','ú°´Áß‡©û ©ú ú°´¨ß‡´„'
   progr=if(p=1,"printbox2","print_xar")
   clear
   sthles=100
sele 1
   USE EIS INDEX HME
   SET SOFTSEEK ON
   SEEK m_eidos + dtos(hm1)
   SET SOFTSEEK Off
   synt3 = 'HME<=HM2 .AND. eidos=m_eidos .and. .not. eof()'

   if m_eidos='*'
      go top
      synt3='.not. eof()'
   endif


if m_taj = '1'
   index on paral to dok for hme>=hm1 .and. hme<=hm2 .and. eidos=m_eidos
endif

if m_taj = '2'
   index on apost to dok for hme>=hm1 .and. hme<=hm2 .and. eidos=m_eidos
endif



IF KOLL <> 1

   do &progr with 9,'&synt3','&synt','000000111',;
       'ÜãÑêéãÜåàÄ','hme','99/99/99',;
       'Äê.èÄêÄíÄí','apa','XXXXXXXXXXXX',;
       'ÄèéëíéäÑÄë','apost','XXXXXXXXXXXXXXX',;
       'èÄêÄäÜèíÜë','paral','XXXXXXXXXXXXXXX',;
       'èéäÜ','polh','XXXXXXXXXX',;
       'ÅÄêéë','baros','###.#',;
       'âÄá.ÄçàÄ','POSO','##,###.##',;
       'èêéãÜáÑàÄ','promhu','##,####.##',;
       'ÄåíàâÄíÄÅ','antikat','##,####.##'

ELSE

   do &progr with 10,'&synt3','&synt','0000000011',;
       'ÜãÑêéãÜåàÄ','hme','99/99/99',;
       'Äê.èÄêÄíÄí','apa','XXXXXXXXXXXX',;
       'ÄèéëíéäÑÄë','apost','XXXXXXXXXXXXXXX',;
       'èÄêÄäÜèíÜë','paral','XXXXXXXXXXXXXXX',;
       'èéäÜ','polh','XXXXXXXXXX',;
       'ÅÄêéë','baros','###.#',;
       'COL','SPACE(2)+COLLECT','XXX',;
       'PPD','SPACE(2)+PPD'    ,'XXX',;
       'âÄá.ÄçàÄ','IF(COLLECT $ "Nnå§",POSO,0)','##,###.##',;
       'ÄåíàâÄíÄÅ','antikat','##,####.##'

ENDIF

   set softseek off

   USE EIS INDEX HME
RETURN




PROC stoix_parad
***************************
CLEAR
STORE DATE() TO HM1,HM2
TYPOS=3
M_apos ='          '
M_paral='          '
M_etairia ='  '
m_eidos='ë'

@ 1 ,10 say 'ÑàëÑêïéãÑåÄ=ë ÑçÑêïéãÑåÄ=ç' get m_eidos
@ 2,10 SAY 'Äèé ' get hm1
@ 4,10 SAY 'Äèé ' get hm2
@ 6,10 say '1=collect 2=ß†©´‡©û 3=¶¢ò' get typos pict '9'
@ 8,10 say  'ÄèéëíéäÑÄë ' get m_apos
@ 10,10 say 'èÄêÄäÜèíÜë ' get m_paral
@ 12,10 say 'ÑíÄàêàÄ    ' get m_etairia
read
SYNT='HME>=HM1 .AND. HME<=HM2'
IF TYPOS=1
   SYNT=SYNT+'.AND. COLLECT $ "Nn§å" '
ENDIF
IF TYPOS=2
   SYNT=SYNT+'.AND. PISTOSH $ "Nn§å" '
ENDIF
IF M_apos <> '  '
   SYNT=SYNT+'.AND. apost=trim(m_apos) '
ENDIF
IF M_paral <> '  '
   SYNT=SYNT+'.AND. paral=trim(m_paral) '
ENDIF
IF M_etairia <> '  '
   SYNT=SYNT+'.AND. etairia = trim(m_etairia) '
ENDIF


 p=1
    do epilogh3 with 12,10,p,2,'ú°´Áß‡©û ©ú ¶üÊ§û','ú°´Áß‡©û ©ú ú°´¨ß‡´„'
   progr=if(p=1,"printbox2","print_xar")
   clear
   sthles=100
sele 1
   USE EIS INDEX HME
   SET SOFTSEEK ON
   SEEK m_eidos + dtos(hm1)

   do &progr with 7,'HME<=HM2 .AND. eidos=m_eidos .and. .not. eof()','&synt','0000000',;
       'ÜãÑêéãÜåàÄ','hme','99/99/99',;
       'Äê.èÄêÄíÄí','apa','XXXXXXXXXXXX',;
       'ÄèéëíéäÑÄë','apost','XXXXXXXXXXXXXXX',;
       'èÄêÄäÜèíÜë','paral','XXXXXXXXXXXXXXX',;
       'ÜãÑê.èÄêÄÉ','HME_PAR','99/99/99',;
       'óêÄ èÄêÄÉ','ORA_PAR','XXXXXXXX',;
       'ëTOIX.èÄêÄÉ','STOIX_PAR','XXXXXXXXXXXXXXX'


   set softseek off
RETURN




PROC ANAZHT
***********************
   CLEAR
   m_eidos='ë'
   m_apa  =space(12)
   m_hme  =date()
   @ 2,2 say 'ÑàëÑêïéãÑåÄ=ë ÑçÑêïéãÑåÄ=ç' get m_eidos
   @ 3,2 say 'Äê.èÄêÄëíÄíàâéì           ' get m_apa
   @ 4,2 say 'ÜãÑêéãÜåàÄ                ' get m_hme
   READ
   USE EIS INDEX APA,HME
   IF m_apa <> '  '
      seek m_eidos + trim ( m_apa )
   else
      set order to 2
      seek m_eidos + dtos(m_hme)
   endif

   do while .t.
      if m_eidos='ë'
         do mask_eis
      else
         do mask_ej
      endif
      clear gets

      @ 24,0 SAY 'Enter=õ†¶®üÈ§‡ Esc=≠úÁö‡ PgDn=úßÊ£ú§û PgUp=ß®¶ûö¶Á£ú§û Delete=õ†òö®ò≠„'
      inkey(0)

      if lastkey()=13  && ENTER
         if m_eidos='ë'
            do mask_eis
         else
            do mask_ej
         endif
         READ
      endif

      if lastkey()=7  && Delete
         ppw=1
         do epilogh3 with 19,60,ppw,2,'úß†©´®‚≠‡','õ†òö®·≠‡'
         if ppw=2
            dele
            wait ' úö†§ú û õ†òö®ò≠„'
            SKIP
         endif
      endif


      if lastkey()=3  && PgDn
         skip
      endif
      if lastkey()=18  && PgUp
         skip -1
      endif



      if lastkey()=27  &&  Esc
         close data
         return
      endif
   enddo

RETURN



PROC EISERX
***********************************************
CLEAR
USE EIS INDEX APA,HME
APPEND BLANK
 REPL EIDOS WITH 'ë'

TEXT
⁄ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒEIëÑêïéãÑåÄƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒø
≥                                                                           ≥
≥ HME   /  /      Äê.èÄêÄëíÄíàâéì  .....                                    ≥
≥                                                                           ≥
≥ ÄèéëíéäÑÄë                                 èÄêÄä                          ≥
≥ èéäÜ                                       BAêOë     .0                   ≥
≥ èÑêàÇêÄîÜ                                                                 ≥
≥                                                                           ≥
≥  PPD                                                                      ≥
≥  COLLECT                                                                  ≥
≥                                                ÜãÑê.èÄêÄä    /  /         ≥
≥                                                óêÄ èÄêÄäÄÅ                ≥
≥  âÄáÄêé                                        ëíéàï.èÄêÄä                ≥
≥    îèÄ                 ANTIKAT                                            ≥
≥  ëìåéäé                                                                   ≥
≥                     MAKêINEë                                              ≥
≥                                                                           ≥
≥  èêéãÜá                                       ÑíÄàêÑàÄ                    ≥
≥                                                                           ≥
≥                                                                           ≥
¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
ENDTEXT

SET COLOR TO &wn,&NW,,,&WPN
@ 3,6 GET HME
@ 3,35 GET APA
@ 5,13 GET APOST
@ 5,52 GET PARAL  valid intpar('pel','kod','epo','paral')=0
@  6,8  GET POLH
@  6,55 GET BAROS
@ 7, 13 GET PERIGRAFH
@ 9,10 GET PPD valid ppd $ 'no§¶åéNO'
READ

IF PPD $ 'o¶éO'
  REPL COLLECT WITH 'N'
  @ 10,10 SAY  COLLECT
  PPW=1
  DO EPILOGH3 WITH 9,21,PPW,2,'èàëíóëÜ','CASH'
  IF ppw=1
     repl pistOSH with 'N'
  ENDIF
  IF PPW=2
     REPL EJOFL WITH 'N'
  ENDIF
ENDIF


@ 13,13 GET POSO  VALID FIND_FPA(0)=0
@ 14,13 GET FPA
@ 15,13 GET synolo


@ 14,36 GET ANTIKAT
@ 18,13 GET PROMHU
@ 18,58 GET ETAIRIA
@ 16,44 GET MAKRINES

@ 11,61 GET hme_par
@ 12,61 GET ora_par
@ 13,61 GET stoix_par

READ

RETURN



PROC EPILOGH3
priv arx_col,pp,i,k,maxlen,row,answer,n,p1,p2,p3,p4
para row,arx_col,answer,n,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11
*** row : ©ú†®· ß·§‡ ò®†©´ú®· ´¶¨ ´ú´®òöÈ§¶¨ arx_col : ©´„¢û ( ò§ -1 °ú§´®·®ú´ò† ©´û§ ö®ò££„ )
*** n:number of questions  answer:variable with the answer 0-n
priv LAST_COLOR

LAST_COLOR = SETCOLOR()
declare pp[15]
for i=1 to n
  k= "p"+alltrim( str(i,2) )
  pp[i] = &k
next
set colo to &NW,&wpn
**save screen
maxlen = len ( pp[1] )

for i = 2 to n       && ............ô®Â©°‡ ´¶ £‚ö†©´¶ £„°¶™
  if len ( pp[i] ) > maxlen
       maxlen = len ( pp[i] )
  endif
next
arx_col = IF ( arx_col = -1 , 40 - maxlen / 2 , arx_col )
@ row,arx_col-1,row+n+1,arx_col+maxlen box "…Õª∫ºÕ»∫"

for i = 1 to n       && .............õÂ§‡ ©‚ Ê¢ò ´ò prompt ´¶ Âõ†¶ £„°¶™
  pp[i] = left ( pp[i] + spac(maxlen) , maxlen )
  @ row+i,arx_col prompt pp[i]
next

menu to answer
SET COLO TO &WN,&NW
SETCOLOR(LAST_COLOR)
return



PROC EJERX
**************************************
 CLEAR
 USE EIS INDEX APA,HME
 APPEND BLANK
 REPL EIDOS WITH 'ç'
TEXT
⁄ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ ÑçÑêïéãÑåÄ ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒø
≥                                                                           ≥
≥ HME 01/01/95    Äê.èÄêÄëíÄíàâéì  12345678                                 ≥
≥                                                                           ≥
≥ ÄèéëíéäÑÄë LAGAKHS                         èÄêÄä  KKFKGKGKFGKFKGFKFKF     ≥
≥ èéäÜ                                       BAêOë        0.0               ≥
≥ èÑêàÇêÄîÜ                                                                 ≥
≥                                                                           ≥
≥  cash                                                                     ≥
≥  ß†©´‡©û                                                                  ≥
≥  collect                                       ÜãÑê.èÄêÄä    /  /         ≥
≥                                                óêÄ èÄêÄäÄÅ                ≥
≥   èéëé                                         ëíéàï.èÄêÄä                ≥
≥    îèÄ                 ANTIKAT                                            ≥
≥ ëìåéäé                                                                    ≥
≥                        MAKêINEë                                           ≥
≥                                                                           ≥
≥ èêéãÜáÑàÄ                               ETAIPIA                           ≥
≥                                                                           ≥
≥                                                                           ≥
¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
ENDTEXT

SET COLOR TO &wn,&NW,,,&WPN
@ 3,6 GET HME
@ 3,35 GET APA
@ 5,13 GET APOST valid intpar('pel','kod','epo','apost')=0
@ 5,52 GET PARAL
@  6,8  GET POLH
@  6,55 GET BAROS
@ 7, 13 GET PERIGRAFH
READ



  PPW=1
  DO EPILOGH3 WITH 8,4,PPW,3,'CASH','èàëíóëÜ','COLLECT'

  DO CASE
     CASE PPW=3
        REPL COLLECT WITH 'N'
     CASE PPW=2
        repl pistOSH with 'N'
     CASE PPW=1
        REPL EJOFL WITH 'N'
  ENDCASE


@ 13,13 GET POSO  VALID FIND_FPA(0)=0
@ 14,13 GET FPA
@ 15,13 GET synolo


@ 14,36 GET ANTIKAT
@ 18,13 GET PROMHU
@ 18,58 GET ETAIRIA
@ 16,44 GET MAKRINES

@ 11,61 GET hme_par
@ 12,61 GET ora_par
@ 13,61 GET stoix_par


READ
RETURN


FUNCTION FIND_FPA
 PARA DUM
 REPL FPA WITH POSO * 0.18
 REPL SYNOLO WITH POSO+FPA
RETURN 0


proc mask_eis
*********************************************************
CLEAR screen
TEXT
⁄ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒEIëÑêïéãÑåÄƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒø
≥                                                                           ≥
≥ HME   /  /      Äê.èÄêÄëíÄíàâéì  .....                                    ≥
≥                                                                           ≥
≥ ÄèéëíéäÑÄë                                 èÄêÄä                          ≥
≥ èéäÜ                                       BAêOë     .0                   ≥
≥ èÑêàÇêÄîÜ                                                                 ≥
≥                                                                           ≥
≥  PPD                                                                      ≥
≥  COLLECT                                                                  ≥
≥  ß†©´‡©û                                       ÜãÑê.èÄêÄä    /  /         ≥
≥  ú•Ê≠¢û©                                       óêÄ èÄêÄäÄÅ                ≥
≥  âÄáÄêé                                        ëíéàï.èÄêÄä                ≥
≥    îèÄ                 ANTIKAT                                            ≥
≥  ëìåéäé                                                                   ≥
≥                     MAKêINEë                                              ≥
≥                                                                           ≥
≥  èêéãÜá                                   ETAIPIA                         ≥
≥                                                                           ≥
≥                                                                           ≥
¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
ENDTEXT

SET COLOR TO &wn,&NW,,,&WPN
@ 3,6 GET HME
@ 3,35 GET APA
@ 5,13 GET APOST
@ 5,52 GET PARAL
@  6,8  GET POLH
@  6,55 GET BAROS
@ 7, 13 GET PERIGRAFH
@ 9,10 GET PPD valid ppd $ 'no§¶åéNO'
@10,10 GET collect
@11,10 GET pistosh
@12,10 GET ejofl



@ 13,13 GET POSO  VALID FIND_FPA(0)=0
@ 14,13 GET FPA
@ 15,13 GET synolo


@ 14,36 GET ANTIKAT
@ 18,13 GET PROMHU
@ 18,58 GET ETAIRIA
@ 16,44 GET MAKRINES

@ 11,61 GET hme_par
@ 12,61 GET ora_par
@ 13,61 GET stoix_par
return


PROC mask_EJ
**************************************
 CLEAR  screen
TEXT
⁄ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ ÑçÑêïéãÑåÄ ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒø
≥                                                                           ≥
≥ HME 01/01/95    Äê.èÄêÄëíÄíàâéì  12345678                                 ≥
≥                                                                           ≥
≥ ÄèéëíéäÑÄë LAGAKHS                         èÄêÄä  KKFKGKGKFGKFKGFKFKF     ≥
≥ èéäÜ                                       BAêOë        0.0               ≥
≥ èÑêàÇêÄîÜ                                                                 ≥
≥                                                                           ≥
≥  cash                                                                     ≥
≥  ß†©´‡©û                                                                  ≥
≥  collect                                       ÜãÑê.èÄêÄä    /  /         ≥
≥                                                óêÄ èÄêÄäÄÅ                ≥
≥   èéëé                                         ëíéàï.èÄêÄä                ≥
≥    îèÄ                 ANTIKAT                                            ≥
≥ ëìåéäé                                                                    ≥
≥                        MAKêINEë                                           ≥
≥                                                                           ≥
≥ èêéãÜáÑàÄ                                  ETAIPIA                        ≥
≥                                                                           ≥
≥                                                                           ≥
¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
ENDTEXT

SET COLOR TO &wn,&NW,,,&WPN
@ 3,6 GET HME
@ 3,35 GET APA
@ 5,13 GET APOST
@ 5,52 GET PARAL
@  6,8  GET POLH
@  6,55 GET BAROS
@ 7, 13 GET PERIGRAFH
@ 9, 13 GET EJOFL
@10, 13 GET pistosh
@11, 13 GET collect


@ 13,13 GET POSO  VALID FIND_FPA(0)=0
@ 14,13 GET FPA
@ 15,13 GET synolo


@ 14,36 GET ANTIKAT
@ 18,13 GET PROMHU
@ 18,58 GET ETAIRIA

@ 16,44 GET MAKRINES

@ 11,61 GET hme_par
@ 12,61 GET ora_par
@ 13,61 GET stoix_par
RETURN

/*
    1  HME         Date           8
    2  APA         Character      8
    3  APOST       Character     25
    4  PARAL       Character     25
    5  POLH        Character     15
    6  POSO        Numeric        5
    7  FPA         Numeric        5
    8  SYNOLO      Numeric        5
    9  PROMHU      Character     15
   10  MAKRINES    Numeric        8
   11  PERIGRAFH   Character     10
   12  BAROS       Numeric        6      1
   13  PPD         Logical        1
   14  COLLECT     Logical        1
   15  EJOFL       Logical        1
   16  PISTOSH     Logical        1
   17  ANTIKAT     Numeric        6
   18  HME_PAR     Date           8
   19  ORA_PAR     Character     10
   20  STOIX_PAR   Character     15
   21  CASH        Logical        1
** Total **                     180

*/



proc PRINTBOX
*********************** P R I N T B O X ************************
**************      °·§ú† ßò®¶¨©Âò©û ò®ÆúÂ¶¨
*** n=ò®†ü£Ê™ fields ß¶¨ ßò®¶¨©†·ù¶§´ò†,synuhkh ©¨§ü„°û ö†ò ´¶ DO WHILE
*** synuhkh2 û ©¨§ü„°û ö†ò ´¶  IF, sum_str ô¢‚ßú† ß¶¨ üò °·§ú† ©¶Á£ú™ ,ß.Æ. "00100" °·§ú† ©¶Á£ú™ ©´¶ 3¶ field
*** sum_pic  ´¶ picture ö†· ´†™ ©¶Á£ú™
*** Ei,Fi,Pi  :Ñß†°ú≠ò¢Âõò ßò®¶¨©Âò©û™,Fields ß¶¨ ßò®¶¨©†·ù¶§´ò†,Picture ßò®¶¨©Âò©û™

para n,synuhkh,synuhkh2,sum_str, e1,f1,p1, e2,f2,p2, e3,f3,p3, e4,f4,p4, e5,f5,p5,  e6,f6,p6, e7,f7,p7, e8,f8,p8
priv sumes,synuhkh,synuhkh2,sum_str,cc,pp,e1,e2,e3,e4,e5,e6,e7,e8,f1,f2,f3,f4,f5,f6,f7,f8,p1,p2,p3,p4,p5,p6,p7,p8,N,I


declare sumes[9]

do case      &&  ß®¶©õ†¶®Âùú† ´¶ £„°¶™ ´‡§ ´¨ßÈ©ú‡§
  case n =3
    lsyn = len ( p1+p2+p3 ) + 4
  case n = 4
    lsyn = len ( p1+p2+p3+p4 ) + 5
  case n = 5
    lsyn = len ( p1+p2+p3+p4+p5 ) + 6
  case n = 6
    lsyn = len ( p1+p2+p3+p4+p5+p6 ) + 7
  case n = 7
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7 ) + 8
  case n = 8
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8 ) + 9
 endcase

 if  n > 2   &&  ß®¶©õ†¶®Âùú† ´†™ ©´„¢ú™ ßò®¶¨©Âò©û™
   c1 = int ( 40 - lsyn / 2 )
   c2 = c1 + len (p1) + 1
   c3 = c2 + len (p2) + 1
   c4 = c3 + len(p3) + 1
   if n > 3
     c5 = c4 + len(p4) + 1
     if n > 4
       c6 = c5 + len(p5) + 1
       if n > 5
          c7 = c6 + len(p6) + 1
          if n > 6
             c8 = c7 + len(p7) + 1
             if n > 7
                c9 = c8 + len(p8) + 1
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
endif  &&  2

do case      &&  ù‡ö®ò≠Âùú† ´ò ´ú´®·ö‡§ò
  case n =3
    do boxaki with n,1,3,23,c1,c2,c3,c4
  case n = 4
    do boxaki with n,1,3,23,c1,c2,c3,c4,c5
  case n = 5
    do boxaki with n,1,3,23,c1,c2,c3,c4,c5,c6
  case n = 6
    do boxaki with n,1,3,23,c1,c2,c3,c4,c5,c6,c7
  case n = 7
    do boxaki with n,1,3,23,c1,c2,c3,c4,c5,c6,c7,c8
  case n = 8
    do boxaki with n,1,3,23,c1,c2,c3,c4,c5,c6,c7,c8,c9
 endcase

 set colo to &WPN
   @ 2,c1+1 say E1
   @ 2,c2+1 say E2
   @ 2,c3+1 say E3
   **** §ú¶ °¶££·´†
   zero = 0
   z = "zero"
   m_f1 = if ( subs(sum_str,1,1)="1" , f1 , z )
   m_f2 = if ( subs(sum_str,2,1)="1" , f2 , z )
   m_f3 = if ( subs(sum_str,3,1)="1" , f3 , z )
   store "zero" to m_f4,m_f5,m_f6,m_f7,m_f8
   *******
   if n > 3
     @ 2,c4+1 say E4
     m_f4 = if ( subs(sum_str,4,1)="1" , f4 , z )

     if n > 4
       @ 2,c5+1 say E5
       m_f5 = if ( subs(sum_str,5,1)="1" , f5 , z )

       if n > 5
          m_f6 = if ( subs(sum_str,6,1)="1" , f6 , z )
          @ 2,c6+1 say E6
          if n > 6
             m_f7 = if ( subs(sum_str,7,1)="1" , f7 , z )
             @ 2,c7+1 say E7
             if n > 7
              m_f8 = if ( subs(sum_str,8,1)="1" , f8 , z )
              @ 2,c8+1 say E8
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
set colo to &WN,&WN
@ 3,0 say ""

for i=1 to 9
  sumes[i] = 0
next

synt1 = if ( type("SYNT1")="U",'.t.',synt1)   && Ê´ò§ ‚®Æú´ò† òß¶ ´n§ òß¶ü„°û ¶®Âùú´ò† ´¶ synt1

do while &synuhkh
 IF &synuhkh2   &&  û õúÁ´ú®û ©¨§ü„°û / ò§ õú§ ¨ß·®Æú† ´Ê´ú .í.

   if &synt1   && û ©¨§ü„°û ´û™ òß¶ü„°û™ ò¢¢†È™ .í.

   @ row()+1,c1+1 say &F1 pict "&P1"
   @ row()  ,c2+1 say &F2 pict "&P2"
   @ row()  ,c3+1 say &F3 pict "&P3"
    sumes[1] = sumes[1] + &m_f1
    sumes[2] = sumes[2] + &m_f2
    sumes[3] = sumes[3] + &m_f3

   if n > 3
     @ row()  ,c4+1 say &F4 pict "&P4"
      sumes[4] = sumes[4] + &m_f4
     if n > 4
       @ row()  ,c5+1 say &F5 pict "&P5"
       sumes[5] = sumes[5] + &m_f5

       if n > 5
          sumes[6] = sumes[6] + &m_f6

          @ row()  ,c6+1 say &F6 pict "&P6"

          if n > 6
             sumes[7] = sumes[7] + &m_f7

             @ row()  ,c7+1 say &F7 pict "&P7"
             if n > 7
                 @ row()  ,c8+1 say &F8 pict "&P8"
                 sumes[8] = sumes[8] + &m_f8
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
  endif  && synt1
ENDIF  &&  û õúÁ´ú®û ©¨§ü„°û synuhkh2
if type('TELOS') <> 'U'  &&   ò§ ¶®Â©´û°ú ´¶ telos
      if recn()=telos  && ò§ ‚≠´ò©ú ©´û§ ´ú¢ú¨´òÂò úöö®ò≠„ , §ò ´¶ °·§ú† eof()
             seek '‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡'
      else
             skip
      endif
else
      skip
endif
if row()=22
   if subs(sum_str,1,1)="1"              && ´¨ßÈ§ú† ´†™ ©¶Á£ú™
      @ 24,1+c1 say sumes[1] pict "&p1"
   endif
   if subs(sum_str,2,1)="1"
      @ 24,1+c2 say sumes[2] pict "&p2"
   endif
   if subs(sum_str,3,1)="1"
      @ 24,1+c3 say sumes[3] pict "&p3"
   endif

   if n > 3
     if subs(sum_str,4,1)="1"
        @ 24,1+c4 say sumes[4] pict "&p4"
     endif
     if n > 4
       if subs(sum_str,5,1)="1"
          @ 24,1+c5 say sumes[5] pict "&p5"
       endif
       if n > 5
          if subs(sum_str,6,1)="1"
             @ 24,1+c6 say sumes[6] pict "&p6"
          endif

          if n > 6
             if subs(sum_str,7,1)="1"
                @ 24,1+c7 say sumes[7] pict "&p7"
             endif
             if n > 7
                if subs(sum_str,8,1)="1"
                    @ 24,1+c8 say sumes[8] pict "&p8"
                endif
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
   @ 24,0 say "< Esc > úß†©´®‚≠‡"

   do nwait
   if lastkey()=27
      return
   endif
   @ 3,0 say ""
endif      &&  row() = 22

enddo


if row() < 22
 do while row()< 22
   @ row()+1,c1+1 say spac(len(p1))
   @ row()  ,c2+1 say spac(len(p2))
   @ row()  ,c3+1 say spac(len(p3))
   if n > 3
     @ row()  ,c4+1 say spac(len(p4))
     if n > 4
       @ row()  ,c5+1 say spac(len(p5))
       if n > 5
          @ row()  ,c6+1 say spac(len(p6))
          if n > 6
             @ row()  ,c7+1 say spac(len(p7))
             if n > 7
               @ row()  ,c8+1 say spac(len(p8))
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
enddo  &&  row()<22

endif && row() < 22

   if subs(sum_str,1,1)="1"              && ´¨ßÈ§ú† ´†™ ©¶Á£ú™
      @ 24,1+c1 say sumes[1] pict "&p1"
   endif
   if subs(sum_str,2,1)="1"
      @ 24,1+c2 say sumes[2] pict "&p2"
   endif
   if subs(sum_str,3,1)="1"
      @ 24,1+c3 say sumes[3] pict "&p3"
   endif

   if n > 3
     if subs(sum_str,4,1)="1"
        @ 24,1+c4 say sumes[4] pict "&p4"
     endif
     if n > 4
       if subs(sum_str,5,1)="1"
          @ 24,1+c5 say sumes[5] pict "&p5"
       endif
       if n > 5
          if subs(sum_str,6,1)="1"
             @ 24,1+c6 say sumes[6] pict "&p6"
          endif

          if n > 6
             if subs(sum_str,7,1)="1"
                @ 24,1+c7 say sumes[7] pict "&p7"
             endif
             if n > 7
                if subs(sum_str,8,1)="1"
                    @ 24,1+c8 say sumes[8] pict "&p8"
                endif
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3

do nwait
return







PROC HOR_SCROLL
****************** éêàÖéåíàé SCROLLING *************************
SET CURSOR OFF
     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)
     for k=3 to 22
        @ k,ARXSTHL SAY left(s[k-2],arxh0) + subs ( S[k-2] , arxh,79-arxh0 )
     next
     @ 23,ARXSTHL SAY left(patos,arxh0) + SUBS(PATOS,ARXH,79-arxh0)
     @ 24,ARXSTHL say left(sumarisma , arxh0) + subs ( sumarisma , arxh , 79-arxh0 )
SET CURSOR ON
RETURN







proc BOXAKI
************* ©Æû£ò´†ùú† ¶®ü¶öÈ§†ò ßò®¶¨©Âò©û™ *****************
PRIV  ariumos,r1,r2,r3,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10
para ariumos,r1,r2,r3,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10
*** ÑèàâÑîÄäàÉÄ
@ r1,c1,r2,c2 box "…Õª∫ºÕ»∫"
@ r1,c2,r2,c3 box "ÀÕª∫ºÕ»∫"



***  âÄíó ãÑêéë
@ r2,c1,r3,c2 box "ÃÕŒ∫ Õ»∫"
@ r2,c2,r3,c3 box "ŒÕŒ∫ Õ ∫"


if ariumos > 2


   @ r1,c3,r2,c4 box "ÀÕª∫ºÕ»∫"
   @ r2,c3,r3,c4 box "ŒÕπ∫ºÕ ∫"

 if ariumos > 3
   @ r1,c4,r2,c5 box "ÀÕª∫ºÕ»∫"
   @ r2,c4,r3,c5 box "ŒÕπ∫ºÕ ∫"

  if ariumos > 4
     @ r1,c5,r2,c6 box "ÀÕª∫ºÕ»∫"
     @ r2,c5,r3,c6 box "ŒÕπ∫ºÕ ∫"
    if ariumos > 5
      @ r1,c6,r2,c7 box "ÀÕª∫ºÕ»∫"
      @ r2,c6,r3,c7 box "ŒÕπ∫ºÕ ∫"
      if ariumos > 6
         @ r1,c7,r2,c8 box "ÀÕª∫ºÕ»∫"
         @ r2,c7,r3,c8 box "ŒÕπ∫ºÕ ∫"
         if ariumos > 7
            @ r1,c8,r2,c9 box "ÀÕª∫ºÕ»∫"
            @ r2,c8,r3,c9 box "ŒÕπ∫ºÕ ∫"
            if ariumos > 8
               @ r1,c9,r2,c10 box "ÀÕª∫ºÕ»∫"
               @ r2,c9,r3,c10 box "ŒÕπ∫ºÕ ∫"
            endif && 8
        endif && 7
      endif  && 6
   endif    && 5
  endif    && 4
 endif    && 3
endif  && 2

RETURN


proc PRINT_Xepik
   ***************** ´¨ßÈ§ú† úß†°ú≠ò¢Âõú™ òß¶ ´¶ print_xar

* if pel5<>'fordd'
*   @ PROW(),PCOL() SAY CHR(27)+"x"   && +"1"   && ÅÄÖÑà íÄ NLQ
* endif
/*
   @ 4,c1+1 say E1
   @  4,c2+1 say E2
   @  4,c3+1 say E3
   if n > 3
     @  4,c4+1 say E4
     if n > 4
       @  4,c5+1 say E5
       if n > 5
          @  4,c6+1 say E6
          if n > 6
             @  4,c7+1 say E7
             if n > 7
                @  4,c8+1 say E8
                if n > 8
                     @  4,c9+1 say E9
                     if n > 9
                        @  4,c10+1 say E10
                        if n > 10
                           @  4,c11+1 say E11
                           if n > 11
                              @  4,c12+1 say E12
                              if n > 12
                                 @  4,c13+1 say E13
                                 if n > 13
                                    @  4,c14+1 say E14
                                    if n > 14
                                       @  4,c15+1 say E15
                                    endif   && 14
                                 endif && 13
                              endif && 12
                           endif && 11
                        endif && 10
                     endif && 9
                endif && 8
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
*/
  @ prow()+1,c1+1 say E1
   @  prow(),c2+1 say E2
   @  prow(),c3+1 say E3
   if n > 3
     @  prow(),c4+1 say E4
     if n > 4
       @  prow(),c5+1 say E5
       if n > 5
          @  prow(),c6+1 say E6
          if n > 6
             @  prow(),c7+1 say E7
             if n > 7
                @  prow(),c8+1 say E8
                if n > 8
                     @  prow(),c9+1 say E9
                     if n > 9
                        @  prow(),c10+1 say E10
                        if n > 10
                           @  prow(),c11+1 say E11
                           if n > 11
                              @  prow(),c12+1 say E12
                              if n > 12
                                 @  prow(),c13+1 say E13
                                 if n > 13
                                    @  prow(),c14+1 say E14
                                    if n > 14
                                       @  prow(),c15+1 say E15
                                    endif   && 14
                                 endif && 13
                              endif && 12
                           endif && 11
                        endif && 10
                     endif && 9
                endif && 8
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3




return




proc PRINT_Xsumes
   ******************************** ´¨ßÈ§ú† ´†™ ©¶Á£ú™ ö†ò ´¶ print_xar ********
   PARA TITLOS
                   *@ PROW(),PCOL() SAY CHR(27)+"x"+"1"   && ÅÄÖÑà íÄ NLQ
   @ PROW()+1,1+C1 SAY TITLOS
   if subs(sum_str,1,1)="1"              && ´¨ßÈ§ú† ´†™ ©¶Á£ú™
      @ PROW(),1+c1 say sumes[1] pict "&p1"
   endif
   if subs(sum_str,2,1)="1"
      @ prow(),1+c2 say sumes[2] pict "&p2"
   endif
   if subs(sum_str,3,1)="1"
      @ prow(),1+c3 say sumes[3] pict "&p3"
   endif

   if n > 3
     if subs(sum_str,4,1)="1"
        @ prow(),1+c4 say sumes[4] pict "&p4"
     endif
     if n > 4
       if subs(sum_str,5,1)="1"
          @ prow(),1+c5 say sumes[5] pict "&p5"
       endif
       if n > 5
          if subs(sum_str,6,1)="1"
             @ prow(),1+c6 say sumes[6] pict "&p6"
          endif

          if n > 6
             if subs(sum_str,7,1)="1"
                @ prow(),1+c7 say sumes[7] pict "&p7"
             endif
             if n > 7
                if subs(sum_str,8,1)="1"
                    @ prow(),1+c8 say sumes[8] pict "&p8"
                endif
                if n > 8
                   if subs(sum_str,9,1)="1"
                     @ prow(),1+c9 say sumes[9] pict "&p9"
                   endif
                   if n > 9
                      if subs(sum_str,10,1)="1"
                         @ prow(),1+c10 say sumes[10] pict "&p10"
                      endif
                      if n > 10
                         if subs(sum_str,11,1)="1"
                            @ prow(),1+c11 say sumes[11] pict "&p11"
                         endif
                         if n > 11
                            if subs(sum_str,12,1)="1"
                               @ prow(),1+c12 say sumes[12] pict "&p12"
                            endif
                            if n > 12
                               if subs(sum_str,13,1)="1"
                                  @ prow(),1+c13 say sumes[13] pict "&p13"
                               endif
                               if n > 13
                                  if subs(sum_str,14,1)="1"
                                     @ prow(),1+c14 say sumes[14] pict "&p14"
                                  endif
                                  if n > 14
                                     if subs(sum_str,15,1)="1"
                                        @ prow(),1+c15 say sumes[15] pict "&p15"
                                     endif
                                  endif && 14
                               endif && 13
                            endif && 12
                         endif && 11
                      endif &&  10
                   endif &&  9
                endif &&  8
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
*   @ PROW()+1,PCOL() SAY CHR(27)+"x"+"0"  && ÅÇÄÖÑà íÄ NLQ
return

proc  pr_SUMselidas
   ******************************** ´¨ßÈ§ú† ´†™ ©¶Á£ú™ SELIDOS ö†ò ´¶ print_xar ********
   PARA TITLOS
                   *@ PROW(),PCOL() SAY CHR(27)+"x"+"1"   && ÅÄÖÑà íÄ NLQ
   @ PROW()+1,1+C1 SAY TITLOS
   if subs(sum_str,1,1)="1"              && ´¨ßÈ§ú† ´†™ ©¶Á£ú™
      @ PROW(),1+c1 say sumes[1]-SUMES0[1] pict "&p1"
   endif
   if subs(sum_str,2,1)="1"
      @ prow(),1+c2 say sumes[2]-SUMES0[2] pict "&p2"
   endif
   if subs(sum_str,3,1)="1"
      @ prow(),1+c3 say sumes[3]-SUMES0[3] pict "&p3"
   endif

   if n > 3
     if subs(sum_str,4,1)="1"
        @ prow(),1+c4 say sumes[4]-SUMES0[4] pict "&p4"
     endif
     if n > 4
       if subs(sum_str,5,1)="1"
          @ prow(),1+c5 say sumes[5]-SUMES0[5] pict "&p5"
       endif
       if n > 5
          if subs(sum_str,6,1)="1"
             @ prow(),1+c6 say sumes[6]-SUMES0[6] pict "&p6"
          endif

          if n > 6
             if subs(sum_str,7,1)="1"
                @ prow(),1+c7 say sumes[7]-SUMES0[7] pict "&p7"
             endif
             if n > 7
                if subs(sum_str,8,1)="1"
                    @ prow(),1+c8 say sumes[8]-SUMES0[8] pict "&p8"
                endif
                if n > 8
                   if subs(sum_str,9,1)="1"
                     @ prow(),1+c9 say sumes[9]-SUMES0[9] pict "&p9"
                   endif
                   if n > 9
                      if subs(sum_str,10,1)="1"
                         @ prow(),1+c10 say sumes[10]-SUMES0[10] pict "&p10"
                      endif
                      if n > 10
                         if subs(sum_str,11,1)="1"
                            @ prow(),1+c11 say sumes[11]-SUMES0[11] pict "&p11"
                         endif
                         if n > 11
                            if subs(sum_str,12,1)="1"
                               @ prow(),1+c12 say sumes[12]-SUMES0[12] pict "&p12"
                            endif
                            if n > 12
                               if subs(sum_str,13,1)="1"
                                  @ prow(),1+c13 say sumes[13]-SUMES0[13] pict "&p13"
                               endif
                               if n > 13
                                  if subs(sum_str,14,1)="1"
                                     @ prow(),1+c14 say sumes[14]-SUMES0[14] pict "&p14"
                                  endif
                                  if n > 14
                                     if subs(sum_str,15,1)="1"
                                        @ prow(),1+c15 say sumes[15]-SUMES0[15] pict "&p15"
                                     endif
                                  endif && 14
                               endif && 13
                            endif && 12
                         endif && 11
                      endif &&  10
                   endif &&  9
                endif &&  8
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
*   @ PROW()+1,PCOL() SAY CHR(27)+"x"+"0"  && ÅÇÄÖÑà íÄ NLQ
return







proc PRINT_XAR
*********************** P R I N T _ X A R ************************
**************      °·§ú† ßò®¶¨©Âò©û ò®ÆúÂ¶¨ ©´¶ Æò®´Â
*** n=ò®†ü£Ê™ fields ß¶¨ ßò®¶¨©†·ù¶§´ò†,synuhkh ©¨§ü„°û ö†ò ´¶ DO WHILE
*** synuhkh2 û ©¨§ü„°û ö†ò ´¶  IF, sum_str ô¢‚ßú† ß¶¨ üò °·§ú† ©¶Á£ú™ ,ß.Æ. "00100" °·§ú† ©¶Á£ú™ ©´¶ 3¶ field
*** sum_pic  ´¶ picture ö†· ´†™ ©¶Á£ú™
*** Ei,Fi,Pi  :Ñß†°ú≠ò¢Âõò ßò®¶¨©Âò©û™,Fields ß¶¨ ßò®¶¨©†·ù¶§´ò†,Picture ßò®¶¨©Âò©û™

para n,synuhkh,synuhkh2,sum_str, e1,f1,p1, e2,f2,p2, e3,f3,p3, e4,f4,p4, e5,f5,p5,  e6,f6,p6, e7,f7,p7, e8,f8,p8,e9,f9,p9,e10,f10,p10,e11,f11,p11,e12,f12,p12,e13,f13,p13,e14,f14,p14,e15,f15,p15
priv sumes,synuhkh,synuhkh2,sum_str,cc,pp
* priv e1,e2,e3,e4,e5,e6,e7,e8,f1,f2,f3,f4,f5,f6,f7,f8,p1,p2,p3,p4,p5,p6,p7,p8,N,I,ayj, e8,f8,p8,e9,f9,p9,e10,f10,p10,e11,f11,p11,e12,f12,p12,e13,f13,p13,e14,f14,p14,e15,f15,p15

priv scr2,SELIDES

* epik=space(50)+dtoc ( date() )

scr2 = savescreen(10,1,12,79)
 @ 10,1 clear to 12,79
 @ 10,1 to 12,79

IF  type("STHLES")="U"
   @ 11,2 say 'Ñß†°ú≠ò¢Âõò ' get epik
   read
endif

restscreen(10,1,12,79,scr2)

if ektypoths<>-1
  do while .not. isprinter()
    do center with  10 ,"Ä§¶†•ú ´¶§ Ñ°´¨ß‡´„ °ò† ß·´û©ú ‚§ò ß¢„°´®¶ „ ´¶ < Esc > ö†ò úß†©´®¶≠„"
    do nwait
    if lastkey()=27
       return
    endif
  enddo

else
  set printer to ekt.txt
endif

****set device to print
set print on
set devi to prin

@ prow(),pcol() say chr(27)+"@"


IF ! type("STHLES")="U"

   IF STHLES>80
      if sthles <= 96
         @ prow(),pcol() say chr(27)+chr(58)  && ö†ò 12 cpi ö®·££ò
      else
         @ prow(),pcol() say chr(27)+chr(15)  && ö†ò ©¨£ß¨°§‡£‚§¶ ö®·££ò
      endif
   ENDIF
ENDIF
declare sumes[18], sumes0[18]

do case      &&  ß®¶©õ†¶®Âùú† ´¶ £„°¶™ ´‡§ ´¨ßÈ©ú‡§
  case n  =3
    lsyn = len ( p1+p2+p3 ) + 4
  case n = 4
    lsyn = len ( p1+p2+p3+p4 ) + 5
  case n = 5
    lsyn = len ( p1+p2+p3+p4+p5 ) + 6
  case n = 6
    lsyn = len ( p1+p2+p3+p4+p5+p6 ) + 7
  case n = 7
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7 ) + 8
  case n = 8
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8 ) + 9
  case n = 9
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9 ) + 10
  case n = 10
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10 ) + 11
  case n = 11
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11 ) + 12
  case n = 12
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11+p12) + 13
  case n = 13
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11+p12+p13 ) + 14
  case n = 14
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11+p12+p13+p14 ) + 15
  case n = 15
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11+p12+p13+p14+p15 ) + 16
 

 endcase


   IF lsyn > 150
      sthles=256
   elseIF lsyn > 80
      if lsyn > 96
         sthles=130
         @ prow(),pcol() say chr(27)+chr(15)  && ö†ò ©¨£ß¨°§‡£‚§¶ ö®·££ò
      else
         sthles=96
         @ prow(),pcol() say chr(27)+chr(58)  && ö†ò 12 cpi ö®·££ò
      endif
   ENDIF


 if  n > 2   &&  ß®¶©õ†¶®Âùú† ´†™ ©´„¢ú™ ßò®¶¨©Âò©û™
   c1 = int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )
   c2 = c1 + len (p1) + 1
   c3 = c2 + len (p2) + 1
   c4 = c3 + len(p3) + 1
   if n > 3
     c5 = c4 + len(p4) + 1
     if n > 4
       c6 = c5 + len(p5) + 1
       if n > 5
          c7 = c6 + len(p6) + 1
          if n > 6
             c8 = c7 + len(p7) + 1
             if n > 7
                c9 = c8 + len(p8) + 1
                if n > 8
                   c10 = c9 + len(p9) + 1
                   if n > 9
                        c11 = c10 + len(p10) + 1
                        if n > 10
                           c12 = c11 + len(p11) + 1
                            if n > 11
                              c13 = c12 + len(p12) + 1
                              if n > 12
                                 c14 = c13 + len(p13) + 1
                                 if n > 13
                                    c15 = c14 + len(p14) + 1
                                    if n > 14
                                       c16 = c15 + len(p15) + 1
                                    endif   && 14
                                 endif    && 13
                              endif  && 12
                           endif && 11
                        endif && 10
                   endif && 9
                endif && 8
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
endif  &&  2

*    @ PROW(),PCOL() SAY CHR(27)+"x"    &&   +"1"   && ÅÄÖÑà íÄ NLQ

 IF ! type("EPIK")="U"
   @ 1,IF(type("STHLES")="U",40,STHLES/2) - LEN(EPIK)/2  SAY EPIK
   @ prow()+1,IF(type("STHLES")="U",40,STHLES/2) - LEN(EPIK)/2  SAY  repl('-',len(epik) )
 ENDIF


   SELIDES=1
   @ PROW()+1,LSYN-2 SAY SELIDES PICT '99'

   do print_xEPIK
   @ PROW()+1,c1+1 SAY repl('-',lsyn-2)

***      @ PROW(),PCOL() SAY CHR(27)+"x"            &&   +"0"  && ÅÇÄÖÑà íÄ NLQ

for i=1 to n
  sumes[i] = 0
  sumes0[i] = 0
next
synt1 = if ( type("SYNT1")="U",'.t.',synt1)   && Ê´ò§ ‚®Æú´ò† òß¶ ´n§ òß¶ü„°û ¶®Âùú´ò† ´¶ synt1
AYJ=0
do while &synuhkh
 IF &synuhkh2   &&  û õúÁ´ú®û ©¨§ü„°û / ò§ õú§ ¨ß·®Æú† ´Ê´ú .í.
  AYJ=AYJ+1
  if &synt1   && û ©¨§ü„°û ´û™ òß¶ü„°û™ ò¢¢†È™ .í.
   @ prow()+1,c1+1 say &F1 pict "&P1"
   @ prow()  ,c2+1 say &F2 pict "&P2"
   @ prow()  ,c3+1 say &F3 pict "&P3"
   if subs(sum_str,1,1)="1"
      sumes[1] = sumes[1] + &F1
   endif
   if subs(sum_str,2,1)="1"
      sumes[2] = sumes[2] + &F2
   endif
   if subs(sum_str,3,1)="1"
      sumes[3] = sumes[3] + &F3
   endif

   if n > 3
     @ prow()  ,c4+1 say &F4 pict "&P4"
     if subs(sum_str,4,1)="1"
        sumes[4] = sumes[4] + &F4
     endif
     if n > 4
       @ prow()  ,c5+1 say &F5 pict "&P5"
       if subs(sum_str,5,1)="1"
          sumes[5] = sumes[5] + &F5
       endif
       if n > 5
          @ prow()  ,c6+1 say &F6 pict "&P6"
          if subs(sum_str,6,1)="1"
            sumes[6] = sumes[6] + &F6
         endif

          if n > 6
             @ prow()  ,c7+1 say &F7 pict "&P7"
             if subs(sum_str,7,1)="1"
                sumes[7] = sumes[7] + &F7
             endif
             if n > 7
                @ prow()  ,c8+1 say &F8 pict "&P8"
                if subs(sum_str,8,1)="1"
                   sumes[8] = sumes[8] + &F8
                endif
                if n > 8
                   @ prow()  ,c9+1 say &F9 pict "&P9"
                   if subs(sum_str,9,1)="1"
                      sumes[9] = sumes[9] + &F9
                   endif
                   if n > 9
                       @ prow()  ,c10+1 say &F10 pict "&P10"
                       if subs(sum_str,10,1)="1"
                          sumes[10] = sumes[10] + &F10
                       endif
                       if n > 10
                          @ prow()  ,c11+1 say &F11 pict "&P11"
                          if subs(sum_str,11,1)="1"
                             sumes[11] = sumes[11] + &F11
                          endif
                         if n > 11
                            @ prow()  ,c12+1 say &F12 pict "&P12"
                            if subs(sum_str,11,1)="1"
                               sumes[12] = sumes[12] + &F12
                            endif
                            if n > 12
                               @ prow()  ,c13+1 say &F13 pict "&P13"
                                 if subs(sum_str,11,1)="1"
                                    sumes[13] = sumes[13] + &F13
                                 endif
                                 if n > 13
                                    @ prow()  ,c14+1 say &F14 pict "&P14"
                                    if subs(sum_str,11,1)="1"
                                       sumes[14] = sumes[14] + &F14
                                    endif
                                    if n > 14
                                       @ prow()  ,c15+1 say &F15 pict "&P15"
                                       if subs(sum_str,11,1)="1"
                                          sumes[15] = sumes[15] + &F15
                                       endif
                                    endif && 14
                                 endif && 13
                            endif && 12
                         endif && 11
                       endif &&  10
                   endif &&  9
                endif &&  8
             endif &&  7
          endif &&  6
       endif  &&  5
     endif  &&  4
   endif  &&  3
  endif && û ©¨§ü„°û synt1
ENDIF  &&  û õúÁ´ú®û ©¨§ü„°û synuhkh2


skip
if prow()=55
   @ 58,1 say ''
   do pr_SUMselidas with 'ëìåéäÄ ëÑäàÉÄë'
   do print_xSUMES  with 'ëìåéäÄ ëÑ ãÑíÄîéêÄ'
   for i=1 to n  && òß¶üû°úÁ‡ ´ò ©Á§¶¢ò ö†ò §ò £ß¶®È §ò ôö·¢‡ £ú´· ´ò ©¨§¶¢ò ©ú¢Âõò™
      sumes0[i] = sumes[i]
   next

   eject
   SELIDES=SELIDES+1
   @ PROW()+1,LSYN-2 SAY SELIDES PICT '99'
   do print_xEPIK
   do print_xSUMES  with 'ëìåéäÄ Äèé ãÑíÄîéêÄ'
endif      &&  prow() = 55

enddo

   @ PROW()+1,c1+1 SAY repl('-',lsyn-2)


   do print_xSUMES
eject


&& òß¶üû°úÁ‡ ´†™ ©¶Á£ú™ ©´ò e1,e2,.. ö†ò §ò ´†™ Æ®û©†£¶ß¶†È
  for k=1 to n
    aa = 'e' + alltrim( str(k) )
    &aa = sumes[k]
  next



set prin off
set devi to scre
set printer to
return





proc PRINTBOX2
*********************** P R I N T B O X ************************
**************      °·§ú† ßò®¶¨©Âò©û ò®ÆúÂ¶¨
*** n=ò®†ü£Ê™ fields ß¶¨ ßò®¶¨©†·ù¶§´ò†,synuhkh ©¨§ü„°û ö†ò ´¶ DO WHILE
*** synuhkh2 û ©¨§ü„°û ö†ò ´¶  IF, sum_str ô¢‚ßú† ß¶¨ üò °·§ú† ©¶Á£ú™ ,ß.Æ. "00100" °·§ú† ©¶Á£ú™ ©´¶ 3¶ field
*** sum_pic  ´¶ picture ö†· ´†™ ©¶Á£ú™
*** Ei,Fi,Pi  :Ñß†°ú≠ò¢Âõò ßò®¶¨©Âò©û™,Fields ß¶¨ ßò®¶¨©†·ù¶§´ò†,Picture ßò®¶¨©Âò©û™

para n,synuhkh,synuhkh2,sum_str, e1,f1,p1, e2,f2,p2, e3,f3,p3, e4,f4,p4, e5,f5,p5,  e6,f6,p6, e7,f7,p7, e8,f8,p8, e9,f9,p9 ,e10,f10,p10,e11,f11,p11, e12,f12,p12,e13,f13,p13,e14,f14,p14,e15,f15,p15
priv sumes,synuhkh,synuhkh2,sum_str,cc,pp
**  ,e1,e2,e3,e4,e5,e6,e7,e8,e9,f1,f2,f3,f4,f5,f6,f7,f8,f9,p1,p2,p3,p4,p5,p6,p7,p8,p9,N,I,S,e10,f10,p10,e11,f11,p11, e12,f12,p12,e13,f13,p13,e14,f14,p14,e15,f15,p15


declare S[25],sumes[15],OUON[5]
for k=1 to 15
   sumes[k]=0
next
for k=1 to 5
   ouon[k]=''
next



do case      &&  ß®¶©õ†¶®Âùú† ´¶ £„°¶™ ´‡§ ´¨ßÈ©ú‡§
  case n =3
    lsyn = len ( p1+p2+p3) + 4
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫"


  case n = 4
    lsyn = len ( p1+p2+p3+p4) + 5
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫"


  case n = 5
    lsyn = len ( p1+p2+p3+p4+p5) + 6
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫"

  case n = 6
    lsyn = len ( p1+p2+p3+p4+p5+p6 ) + 7
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"


  case n = 7
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7 ) + 8
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"

  case n = 8
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8 ) + 9
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"À"+repl("Õ",len(p8))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"+left(e8+space(50),len(p8))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"Œ"+repl("Õ",len(p8))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+" "+repl("Õ",len(p8))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"+left(space(50),len(p8))+"∫"

  case n=9
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9 ) + 10
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"À"+repl("Õ",len(p8))+"À"+repl("Õ",len(p9))+ "ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"+left(e8+space(50),len(p8))+"∫"+left(e9+space(50),len(p9))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"Œ"+repl("Õ",len(p8))+"Œ"+repl("Õ",len(p9))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+" "+repl("Õ",len(p8))+" "+repl("Õ",len(p9))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"+left(space(50),len(p8))+"∫"+left(space(50),len(p9))+"∫"
  case n=10
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10 ) + 11
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"À"+repl("Õ",len(p8))+"À"+repl("Õ",len(p9))+"À"+repl("Õ",len(p10))+ "ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"+left(e8+space(50),len(p8))+"∫"+left(e9+space(50),len(p9))+"∫" +left(e10+space(50),len(p10))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"Œ"+repl("Õ",len(p8))+"Œ"+repl("Õ",len(p9))+"Œ"+repl("Õ",len(p10))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+" "+repl("Õ",len(p8))+" "+repl("Õ",len(p9))+" "+repl("Õ",len(p10))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"+left(space(50),len(p8))+"∫"+left(space(50),len(p9))+"∫"+left(space(50),len(p10))+"∫"

  case n=11
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11 ) + 12
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"À"+repl("Õ",len(p8))+"À"+repl("Õ",len(p9))+"À"+repl("Õ",len(p10))+"À"+repl("Õ",len(p11))+ "ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"+left(e8+space(50),len(p8))+"∫"+left(e9+space(50),len(p9))+"∫" +left(e10+space(50),len(p10))+"∫"+left(e11+space(50),len(p11))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"Œ"+repl("Õ",len(p8))+"Œ"+repl("Õ",len(p9))+"Œ"+repl("Õ",len(p10))+"Œ"+repl("Õ",len(p11))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+" "+repl("Õ",len(p8))+" "+repl("Õ",len(p9))+" "+repl("Õ",len(p10))+" "+repl("Õ",len(p11))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"+left(space(50),len(p8))+"∫"+left(space(50),len(p9))+"∫"+left(space(50),len(p10))+"∫"+left(space(50),len(p11))+"∫"
  


  case n=12
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11+p12 ) + 13
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"À"+repl("Õ",len(p8))+"À"+repl("Õ",len(p9))+"À"+repl("Õ",len(p10))+"À"+repl("Õ",len(p11))+"À"+repl("Õ",len(p12))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"+left(e8+space(50),len(p8))+"∫"+left(e9+space(50),len(p9))+"∫" +left(e10+space(50),len(p10))+"∫"+left(e11+space(50),len(p11))+"∫"+left(e12+space(50),len(p12))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"Œ"+repl("Õ",len(p8))+"Œ"+repl("Õ",len(p9))+"Œ"+repl("Õ",len(p10))+"Œ"+repl("Õ",len(p11))+"Œ"+repl("Õ",len(p12))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+" "+repl("Õ",len(p8))+" "+repl("Õ",len(p9))+" "+repl("Õ",len(p10))+" "+repl("Õ",len(p11))+" "+repl("Õ",len(p12))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"+left(space(50),len(p8))+"∫"+left(space(50),len(p9))+"∫"+left(space(50),len(p10))+"∫"+left(space(50),len(p11))+"∫"+left(space(50),len(p12))+"∫"



  case n=13
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11+p12+p13) + 14
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"À"+repl("Õ",len(p8))+"À"+repl("Õ",len(p9))+"À"+repl("Õ",len(p10))+"À"+repl("Õ",len(p11))+"À"+repl("Õ",len(p12))+"À"+repl("Õ",len(p13))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"+left(e8+space(50),len(p8))+"∫"+left(e9+space(50),len(p9))+"∫" +left(e10+space(50),len(p10))+"∫"+left(e11+space(50),len(p11))+"∫"+left(e12+space(50),len(p12))+"∫"+left(e13+space(50),len(p13))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"Œ"+repl("Õ",len(p8))+"Œ"+repl("Õ",len(p9))+"Œ"+repl("Õ",len(p10))+"Œ"+repl("Õ",len(p11))+"Œ"+repl("Õ",len(p12))+"Œ"+repl("Õ",len(p13))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+" "+repl("Õ",len(p8))+" "+repl("Õ",len(p9))+" "+repl("Õ",len(p10))+" "+repl("Õ",len(p11))+" "+repl("Õ",len(p12))+" "+repl("Õ",len(p13))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"+left(space(50),len(p8))+"∫"+left(space(50),len(p9))+"∫"+left(space(50),len(p10))+"∫"+left(space(50),len(p11))+"∫"+left(space(50),len(p12))+"∫"+left(space(50),len(p13))+"∫"


  case n=14
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11+p12+p13+p14 ) + 15
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"À"+repl("Õ",len(p8))+"À"+repl("Õ",len(p9))+"À"+repl("Õ",len(p10))+"À"+repl("Õ",len(p11))+"À"+repl("Õ",len(p12))+"À"+repl("Õ",len(p13))+"À"+repl("Õ",len(p14))+ "ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"+left(e8+space(50),len(p8))+"∫"+left(e9+space(50),len(p9))+"∫" +left(e10+space(50),len(p10))+"∫"+left(e11+space(50),len(p11))+"∫"+left(e12+space(50),len(p12))+"∫"+left(e13+space(50),len(p13))+"∫"+left(e14+space(50),len(p14))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"Œ"+repl("Õ",len(p8))+"Œ"+repl("Õ",len(p9))+"Œ"+repl("Õ",len(p10))+"Œ"+repl("Õ",len(p11))+"Œ"+repl("Õ",len(p12))+"Œ"+repl("Õ",len(p13))+"Œ"+repl("Õ",len(p14))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+" "+repl("Õ",len(p8))+" "+repl("Õ",len(p9))+" "+repl("Õ",len(p10))+" "+repl("Õ",len(p11))+" "+repl("Õ",len(p12))+" "+repl("Õ",len(p13))+" "+repl("Õ",len(p14))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"+left(space(50),len(p8))+"∫"+left(space(50),len(p9))+"∫"+left(space(50),len(p10))+"∫"+left(space(50),len(p11))+"∫"+left(space(50),len(p12))+"∫"+left(space(50),len(p13))+"∫"+left(space(50),len(p14))+"∫"

  case n=15
    lsyn = len ( p1+p2+p3+p4+p5+p6+p7+p8+p9+p10+p11+p12+p13+p14+p15 ) + 16
    epikef0 = "…"+repl("Õ",len(p1))+"À"+repl("Õ",len(p2))+"À"+repl("Õ",len(p3))+"À"+repl("Õ",len(p4))+"À"+repl("Õ",len(p5))+"À"+repl("Õ",len(p6))+"À"+repl("Õ",len(p7))+"À"+repl("Õ",len(p8))+"À"+repl("Õ",len(p9))+"À"+repl("Õ",len(p10))+"À"+repl("Õ",len(p11))+"À"+repl("Õ",len(p12))+"À"+repl("Õ",len(p13))+"À"+repl("Õ",len(p14))+ "À"+repl("Õ",len(p15))+"ª"
    epikef = "∫"+  left(e1+space(50),len(p1)) + "∫"+  left(e2+space(50),len(p2)) +"∫"+  left(e3+space(50),len(p3)) + "∫" +  left(e4+space(50),len(p4)) + "∫" ;
        +  left(e5+space(50),len(p5)) + "∫" +  left(e6+space(50),len(p6))+"∫"+left(e7+space(50),len(p7))+"∫"+left(e8+space(50),len(p8))+"∫"+left(e9+space(50),len(p9))+"∫" +left(e10+space(50),len(p10))+"∫"+left(e11+space(50),len(p11))+"∫"+left(e12+space(50),len(p12))+"∫"+left(e13+space(50),len(p13))+"∫"+left(e14+space(50),len(p14))+"∫"+left(e15+space(50),len(p15))+"∫"
    epikef2 = "Ã"+repl("Õ",len(p1))+"Œ"+repl("Õ",len(p2))+"Œ"+repl("Õ",len(p3))+"Œ"+repl("Õ",len(p4))+"Œ"+repl("Õ",len(p5))+"Œ"+repl("Õ",len(p6))+"Œ"+repl("Õ",len(p7))+"Œ"+repl("Õ",len(p8))+"Œ"+repl("Õ",len(p9))+"Œ"+repl("Õ",len(p10))+"Œ"+repl("Õ",len(p11))+"Œ"+repl("Õ",len(p12))+"Œ"+repl("Õ",len(p13))+"Œ"+repl("Õ",len(p14))+"Œ"+repl("Õ",len(p15))+"π"
    PATOS = "»"+repl("Õ",len(p1))+" "+repl("Õ",len(p2))+" "+repl("Õ",len(p3))+" "+repl("Õ",len(p4))+" "+repl("Õ",len(p5))+" "+repl("Õ",len(p6))+" "+repl("Õ",len(p7))+" "+repl("Õ",len(p8))+" "+repl("Õ",len(p9))+" "+repl("Õ",len(p10))+" "+repl("Õ",len(p11))+" "+repl("Õ",len(p12))+" "+repl("Õ",len(p13))+" "+repl("Õ",len(p14))+" "+repl("Õ",len(p15))+"º"
    GEMIS = "∫"+  left(space(50),len(p1)) + "∫"+  left(space(50),len(p2)) +"∫"+  left(space(50),len(p3)) + "∫" +  left(space(50),len(p4)) + "∫" ;
        +  left(space(50),len(p5)) + "∫" +  left(space(50),len(p6))+"∫"+left(space(50),len(p7))+"∫"+left(space(50),len(p8))+"∫"+left(space(50),len(p9))+"∫"+left(space(50),len(p10))+"∫"+left(space(50),len(p11))+"∫"+left(space(50),len(p12))+"∫"+left(space(50),len(p13))+"∫"+left(space(50),len(p14))+"∫"+left(space(50),len(p15))+"∫"
 endcase
 do printbox21
return



proc printbox21
****************** ©¨§‚Æú†ò òß¶ ßò®òß·§‡ ***********************************
synt1 = if ( type("SYNT1")="U",'.t.',synt1)   && Ê´ò§ ‚®Æú´ò† òß¶ ´n§ òß¶ü„°û ¶®Âùú´ò† ´¶ synt1

SEIRA = 1




CLEAR

arxh0=1+len(p1+p2)+2
arxh = arxh0 + 1
nouon=0

ARXSTHL = IF ( LSYN>77 , 0 , 40 - LSYN/2 )

     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)


do while &synuhkh
     set cursor off


 IF &synuhkh2   &&  û õúÁ´ú®û ©¨§ü„°û / ò§ õú§ ¨ß·®Æú† ´Ê´ú .í.

   if &synt1   && û ©¨§ü„°û ´û™ òß¶ü„°û™ ò¢¢†È™ .í.
      do case      &&  ß®¶©õ†¶®Âùú† ´¶ £„°¶™ ´‡§ ´¨ßÈ©ú‡§
          case n =3
            u="∫"+ trans(&f1,"&p1") +"∫"+ trans(&f2,"&p2") +"∫"+ trans(&f3,"&p3")+"∫"
          case n = 4
            u="∫"+  trans(&f1,"&p1") +"∫"+ trans(&f2,"&p2") +"∫"+ trans(&f3,"&p3") +"∫"+ trans(&f4,"&p4")+"∫"
          case n = 5
            u="∫"+  trans(&f1,"&p1") +"∫"+ trans(&f2,"&p2") +"∫"+ trans(&f3,"&p3") +"∫"+ trans(&f4,"&p4") +"∫"+ trans(&f5,"&p5")+"∫"
          case n = 6
            u= "∫"+ trans(&f1,"&p1") +"∫"+ trans(&f2,"&p2") +"∫"+ trans(&f3,"&p3") +"∫"+ trans(&f4,"&p4") +"∫"+ trans(&f5,"&p5") +"∫";
                + trans(&f6,"&p6")+"∫"
          case n = 7
            u= "∫"+ trans(&f1,"&p1") +"∫"+ trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+   trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +;
               "∫"+trans(&f6,"&p6") + "∫"+trans(&f7,"&p7")+"∫"
          case n = 8
            u= "∫"+trans(&f1,"&p1") + "∫"+trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +"∫";
                +trans(&f6,"&p6")+"∫"+trans(&f7,"&p7")+"∫"+trans(&f8,"&p8")+"∫"
          case n=9
            u= "∫"+trans(&f1,"&p1") + "∫"+trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +"∫";
                +trans(&f6,"&p6")+"∫"+trans(&f7,"&p7")+"∫"+trans(&f8,"&p8")+"∫"+trans(&f9,"&p9")+"∫"
          case n=10
            u= "∫"+trans(&f1,"&p1") + "∫"+trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +"∫";
                +trans(&f6,"&p6")+"∫"+trans(&f7,"&p7")+"∫"+trans(&f8,"&p8")+"∫"+trans(&f9,"&p9")+"∫"+trans(&f10,"&p10")+"∫"
          case n=11
            u= "∫"+trans(&f1,"&p1") + "∫"+trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +"∫";
                +trans(&f6,"&p6")+"∫"+trans(&f7,"&p7")+"∫"+trans(&f8,"&p8")+"∫"+trans(&f9,"&p9")+"∫"+trans(&f10,"&p10")+"∫"+trans(&f11,"&p11")+"∫"
          case n=12
            u= "∫"+trans(&f1,"&p1") + "∫"+trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +"∫";
                +trans(&f6,"&p6")+"∫"+trans(&f7,"&p7")+"∫"+trans(&f8,"&p8")+"∫"+trans(&f9,"&p9")+"∫"+trans(&f10,"&p10")+"∫"+trans(&f11,"&p11")+"∫"+;
                +trans(&f12,"&p12")+"∫"
          case n=13
            u= "∫"+trans(&f1,"&p1") + "∫"+trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +"∫";
                +trans(&f6,"&p6")+"∫"+trans(&f7,"&p7")+"∫"+trans(&f8,"&p8")+"∫"+trans(&f9,"&p9")+"∫"+trans(&f10,"&p10")+"∫"+trans(&f11,"&p11")+"∫"+;
                +trans(&f12,"&p12")+"∫" +trans(&f13,"&p13")+"∫"
          case n=14
            u= "∫"+trans(&f1,"&p1") + "∫"+trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +"∫";
                +trans(&f6,"&p6")+"∫"+trans(&f7,"&p7")+"∫"+trans(&f8,"&p8")+"∫"+trans(&f9,"&p9")+"∫"+trans(&f10,"&p10")+"∫"+trans(&f11,"&p11")+"∫"+;
                +trans(&f12,"&p12")+"∫" +trans(&f13,"&p13")+"∫"  +trans(&f14,"&p14")+"∫"
          case n=15
            u= "∫"+trans(&f1,"&p1") + "∫"+trans(&f2,"&p2") + "∫"+trans(&f3,"&p3")+"∫"+trans(&f4,"&p4") + "∫"+trans(&f5,"&p5") +"∫";
                +trans(&f6,"&p6")+"∫"+trans(&f7,"&p7")+"∫"+trans(&f8,"&p8")+"∫"+trans(&f9,"&p9")+"∫"+trans(&f10,"&p10")+"∫"+trans(&f11,"&p11")+"∫"+;
                +trans(&f12,"&p12")+"∫" +trans(&f13,"&p13")+"∫"  +trans(&f14,"&p14")+"∫" +trans(&f15,"&p15")+"∫"

         endcase

      S[SEIRA] = u
      @ seira+2,ARXSTHL SAY left(s[seira],arxh0) + subs ( S[SEIRA] , arxh,79-arxh0 )
      SEIRA=SEIRA + 1
      for k=1 to n
          ww = alltrim ( str(k) )
          aas = f&ww
          sumes[k] = sumes[k] +  if ( subs(sum_str,k,1)="1" , &aas  , 0 )
      next
    endif  && synt1
 ENDIF  &&  û õúÁ´ú®û ©¨§ü„°û synuhkh2

if type('TELOS') <> 'U'  &&   ò§ ¶®Â©´û°ú ´¶ telos
      if recn()=telos  && ò§ ‚≠´ò©ú ©´û§ ´ú¢ú¨´òÂò úöö®ò≠„ , §ò ´¶ °·§ú† eof()
           seek '‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡'
      else
           skip
      endif
else
      skip
endif

if ROW() >= 22
 @ ROW()+1,ARXSTHL SAY left(patos,arxh0)+SUBS(PATOS,ARXH,79-arxh0)
 set cursor on
 sumarisma = ''
 for k=1 to n
        ww = alltrim ( str(k) )
        aas = p&ww
        sumarisma = sumarisma + " " +  if ( subs(sum_str,k,1)="1" , trans(sumes[k],"&aas"),space( len(aas) )  )
 next
 sumarisma = sumarisma + ' '
 @ 24,ARXSTHL say left(sumarisma,arxh0)+subs ( sumarisma , arxh , 79-arxh0 )

 nouon = nouon + 1
 if nouon > 5
    for k=1 to 4
        ouon[k] = ouon[k+1]
    next
    save screen to ouon[5]
    nouon=5
 else
    save screen to ouon[nouon]
 endif

 @ 24,ARXSTHL say "< Esc > úß†©´®‚≠‡"
 DO Nwait
 plhk=lastkey()
 piso=0
 do while (plhk=4 .or. plhk=19  .or. plhk=18 .or. plhk=3) .AND. LSYN-n>79 && -> „ <- „ PgDn „ PgUp
   if plhk=18 && PgUp
      piso=min(piso+1,nouon-1)
      restore screen from ouon[nouon-piso]
   endif
   if plhk=3 && PgDn
      piso=max(piso-1,0)
      restore screen from ouon[nouon-piso]
   endif
   if plhk=4 && -> belos
     arxh= min ( arxh+2 ,lsyn-78+arxh0)
     DO HOR_SCROLL
   endif
   if plhk=19 && <- belos
     arxh=max(arxh-2,arxh0+1)
     DO HOR_SCROLL
   endif
   do nwait
   plhk=lastkey()
 enddo
   CLEAR
   SEIRA=1
   if lastkey()=27
      return
   endif
endif      &&  row() = 22

     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)



enddo

 DO WHILE ROW()<22
     S[SEIRA] = GEMIS
     @ seira+2,ARXSTHL SAY left(s[seira],arxh0) + subs ( S[SEIRA] , arxh,79-arxh0 )
     SEIRA= SEIRA + 1
 ENDDO

 @ ROW()+1,ARXSTHL SAY left(patos,arxh0)+SUBS(PATOS,ARXH,79-arxh0)
 set cursor on
 sumarisma = ''
 for k=1 to n
        ww = alltrim ( str(k) )
        aas = p&ww
        sumarisma = sumarisma + " " +  if ( subs(sum_str,k,1)="1" , trans(sumes[k],"&aas"),space( len(aas) )  )
 next
 sumarisma = sumarisma + ' '
 @ 24,ARXSTHL say left(sumarisma,arxh0)+subs ( sumarisma , arxh , 79-arxh0 )

 nouon = nouon + 1
 if nouon > 5
    for k=1 to 4
        ouon[k] = ouon[k+1]
    next
    save screen to ouon[5]
    nouon=5
 else
    save screen to ouon[nouon]
 endif

 @ 24,ARXSTHL say "< Esc > úß†©´®‚≠‡"
 DO Nwait
 plhk=lastkey()
 piso=0
 do while plhk=4 .or. plhk=19  .or. plhk=18 .or. plhk=3 && -> „ <- „ PgDn „ PgUp
   if plhk=18 && PgUp
      piso=min(piso+1,nouon-1)
      restore screen from ouon[nouon-piso]
   endif
   if plhk=3 && PgDn
      piso=max(piso-1,0)
      restore screen from ouon[nouon-piso]
   endif
   if plhk=4 && -> belos
     arxh= min ( arxh+2 ,lsyn-78+arxh0)
     DO HOR_SCROLL
   endif
   if plhk=19 && <- belos
     arxh=max(arxh-2,arxh0+1)
     DO HOR_SCROLL
   endif
   do nwait
   plhk=lastkey()
 enddo
   CLEAR
   SEIRA=1
   if lastkey()=27
      return
   endif
return

proc nwait
**********  waiting untill key pressed *************
a=inkey()
do while a=0
  *@ 23,78 SAY ""
  a=inkey()
enddo
return

proc center
priv row,string
para row,string
******************* °ú§´®·®†©£ò string *****************
@ row,int(40-len(string)/2) say string
return




proc intpar
******************* ßò®¶¨©Âò©û ú†õÈ§ *****************
*cl5 local GetList := {}
local GetList := {}
priv k,i,kk,po,choice,l , inter,kmod,mono,kkk,helplevel,XRVMA,proto_record[15],selida,ssa
priv sslejh,filtro,moo:=space(1),oo:=space(10),arx_kod:=space(10)
declare fp[5]
  fp[1]=1
  fp[2]=1
  fp[3]=1
  fp[4]=1
  fp[5]=1

helplevel=45
para workspace,mkod,mono,m_apost
xrvma = setcolor()
inter = 1
save screen
kkk=str(select(),2)

set colo to &NW,&WPN &&  N/BG,R+/N
declare po [ 15 ] , onoma [ 15 ]
choice=1
@ 9,3 CLEAR TO 24,79
@ 9,3 to 24,79 double
@ 24,20 say 'PgDn úßÊ£ú§û ©ú¢Âõò-PgUp ß®¶ûö¶Á£ú§û ©ú¢Âõò'
sele 2   &&   &workspace
DO WHILE .T.     && ----------------------------------------
  oo     =oo     +space( 14-len(oo) )  && ö†ò´† ò§ ßò´„©‡ ú§´ú® £ú´· õú§ õ‚Æú´ò† ´Âß¶´ò
  oo = if (moo=space(1) , oo , moo ) && ò§ ‚®Æú´ò† ß®È´û ≠¶®ò ( moo=space(1) ) ßò†®§ú† ´ò °ú§· , ò¢¢¶†È™ ßòÂ®§ú† ´û§ ßò¢†· ´†£„ moo
  arx_kod=arx_kod+space( 14-len(arx_kod) )
  ***  ¶®†ùú´ò† public  store "        " to oo,arx_kod
  set colo to &WN,&WPN  && R/n,RB+/n
  @ 10,5 SAY "íé  éåéãÄ åÄ ÄêïàÖÑà Äèé"
  @ 12,5 say "é âóÉàâéë åÄ ÄêïàÖÑà Äèé"
  if updating <> 99   && ú†©òö‡öû ´†£¶¢¶ö†¶¨ £ú¢†©©ò ( ò§´† ö†ò °‡õ†°Ê õÂ§‡ ¶§¶£ò °ò† ´¶ Ø·Æ§ú† òß ú¨üúÂò™ )
     @ 10,30 get oo   && oo,arx_kod = public
     @ 12,30 get arx_kod
     read

     *oo = input_chars ( 10 , 30 , oo )   && oo,arx_kod = public
     *arx_kod = input_chars ( 12 , 30 , arx_kod )
  else  && updating=99
    updating = 1
  endif
  moo=oo
  set colo to   &NW,&WPN
  @ 11,4 CLEAR TO 23,78
  IF LASTKEY()=27
    sele &kkk
    restore screen
    KEYBOARD ""
     setcolor(xrvma  )
    return 0
  endif
  KEYBOARD ""
  ** ÇàÄ åÄ ÅÄÖÑà âÄà îàäíêé   è.ï.    ëÄãè\íàã  éäÄ íÄ ëÄãèéìÄå èéì èÑêàÑïéìå íÜå äÑçÜ íàã
  ssa = at ( '\' , oo )  && ©ú ß¶†· ü‚©û ô®„°ú ´û§ \
  if ssa = 0
        oo = trim (oo)   && ò§ õú§ ‚Æú† ò§·ß¶õû °·üú´¶
        filtro='.t.'
  else
        sslejh =   trim (  subs ( oo , ssa + 1 , len ( oo) - ssa  ) )  && ò¨´Ê ß¶¨ üò Ø·Æ§ú†  (ëÄã\íàã úÂ§ò† ´¶ íàã )
        oo=  left ( oo , ssa - 1 )
        filtro = ' sslejh $ &mono '
  endif
  if ',' $ oo
      if upper(workspace) = "PEL"
         filtro = filtro + ' .and.  typ>0 '
      else
         filtro = filtro + ' .and.  pos>0 '
      endif

      OO = LEFT(OO, at ( ',' , oo )-1  )  && ©ú ß¶†· ü‚©û ô®„°ú ´û§ ","
  endif


  l=len(oo)
  arx_kod= trim(arx_kod)
  ll=len(arx_kod)
  if  workspace = "pel"
    seek buff
    if eof() && õú§ ¨ß·®Æ¶¨§ °òüÊ¢¶¨ ßú¢·´ú™ „ ß®¶£ûüú¨´‚™
      sele &kkk
      restore screen
      KEYBOARD ""
      setcolor(xrvma  )
      return 0
    endif
    if  len(OO)>0  && ò§ ‚õ‡©ú ONOMA
      set order to 2
      seek buff + oo
    else
      set order to 1
      seek buff + arx_kod
    endif
    f3='TRANS(thl,"XXXXXXXXXXX")'
    f4='LEFT(DIE,15)'
    *  f4='DTOC(LASTUPDT)'
  else  &&   if  workspace = "eid"

    IF TPmeFPA=1             &&   PEL5 $ 'melis fordd'
       f3='TRANS(if(fpa=0,0,fp[fpa]*LTI),"9,999,999.9")'
    else
       f3='TRANS(LTI,"9,999,999.9")'
    endif

    f4='TRANS(pos,"99,999.9")'
    if len(oo)>0
      set inde to EIDONO,EIDKO
      seek oo
    else
      set inde to EIDKO,eidono
      seek arx_kod
    endif
  endif
  if eof()
    go top
  endif
  i=0
  selida=1
  proto_record[selida]=recn()
  @ 9,0 say ""
  do while .not. eof()
   if left(&mono,l)=oo  .AND.  left ( &mkod ,ll ) = arx_kod .and. &filtro
     i=i+1
     onoma[i] = left ( &mkod+" "+&mono+' '+&f3 + ' '+ &f4 ,74)
     po [i]= recn()
     @ 9+i,4 prompt subs("123456789ÄÅÇÉÑ",i,1) +". "+ onoma [i]
    if i=14
     menu to choice
     if lastkey()=18   && ò§ ß·´û©ú PgUp (ß®¶ûö¶Á£ú§û ©ú¢Âõò )
         do intpar_2
         loop
     endif  && pGuP

     if lastkey() = 3    && ò§ ß·´û©ú / PgDn (úßÊ£ú§û ©ú¢Âõò )
        i = 0
        proto_record[selida]=po[1]  && °®ò´ò ´¶ ß®È´¶ recored ö†ò §ò Æ®û©†£¶ß¶†ûüúÂ òß¶ to PgUp
        selida=if(selida=15,15,selida+1)
        choice=1
        @ 11,4 CLEAR TO 23,78
        loop
     endif  && PgDn
     exit
   endif    && i=14
  endif       &&  left(&mono,l)=oo
  skip
  if workspace="pel"
     if buff<>eidos
       menu to choice
       if lastkey()=18   && ò§ ß·´û©ú PgUp (ß®¶ûö¶Á£ú§û ©ú¢Âõò )
         do intpar_2
         loop
       endif  && pGuP
       go top
       exit
     endif && buff <> eidos
  endif  && "pel"

    &&------ ll=len(mkod)     l=len(oo)
    && ú·§ ( õú§ ‚õ‡©ú Ê§¶£ò .°ò†. ô®Â©°‡ õ†ò≠¶®ú´†°Ê °‡õ†°Ê òßÊ ´¶§ ò®Æ†°Ê )
    && .„. ( ‚õ‡©ú Ê§¶£ò .°ò†. ô®Â©°‡ Ê§¶£ò õ†ò≠¶®ú´†°Ê òß¶ ´¶ ò®Æ†°Ê )
    && ´Ê´ú ´ú¢ú†È©ò£ú , °ò† õ†ò¢‚öú†™ ò¨´Ê ß¶¨  ü‚¢ú†™

  if  ( l=0 .and. left ( &mkod ,ll ) <> arx_kod )   .or.  ( l>0 .and. left(&mono,l)<>oo )
      menu to choice
      if lastkey()=18   && ò§ ß·´û©ú PgUp (ß®¶ûö¶Á£ú§û ©ú¢Âõò )
         do intpar_2
         loop
      endif  && pGuP
      go top     && ß®Ê©üú´¶
      exit
  endif
enddo  &&  -----------
if i=0       && IF eof() and i=0
  LOOP
endif
EXIT
ENDDO   && .T. ---------------------------
if eof()
  menu to choice
endif
set colo to &WN,&WN
*** neo
if upper(workspace)="PEL" .AND. upper ( indexkey(0) ) <>'EIDOS+KOD'   && len(oo) > 0
   set order to 1
endif
if upper(workspace)="EID" .AND.  upper ( indexkey(0) ) <>'KOD'  && len(oo) > 0
   set order to 1
endif


if choice <> 0 .and. choice<> i+1 .and. lastkey()<>27
  go po [ choice ]
endif
sele &kkk
restore screen
KEYBOARD ""
  setcolor(xrvma  )
  oo = moo

if left(m_apost,2)='m_'
   &m_apost  = pel->epo
else
  repl &m_apost with pel->epo
  if m_apost = 'apost'
      @ 5,13 say APOST
  else
      @ 5,52 say paral
  endif
  @ 5,13 say APOST
endif


return 0


proc intpar_2
**************** ´† öÂ§ú´ò† Ê´ò§ ßò´· PgDn ***********
         i = 0
         choice=1
         @ 11,4 CLEAR TO 23,78
         selida=if(selida=1,1,selida-1)
         go proto_record[selida]
return


