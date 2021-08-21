Attribute VB_Name = "Module5"
Option Explicit
Dim f_kodik As Recordset, F_T As Recordset, f_tab(40)
Dim f_db As Database





Function print4_xar(sql As String, SUgm_str, EPIKEF As String, GROUPN As Integer)
'"*********************** P R I N T _ X A R ************************
'"**************      κάνει παρουσίαση αρχείου στο χαρτί
'"*** n=αριθμός fields που παρουσιάζονται
'"*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
'"*** sum_pic  το picture γιά τις σούμες
'"*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
'"** synt_eject:=0  αλλάζει σελίδα όταν μεταβάλλεται η στήλη
Dim mfields(120), mSYN
Dim synt_eject
Dim SUMES, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), f(120)
Dim mBSEIRA
Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE
Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4
Dim aaP, aaP2, ektypoths
Dim EPIK, xeirisths, PPF, epik3
Dim sthles, kw, PPD, SELIS
Dim LSYN, aa, epik2, I As Integer
Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject
Dim m_sumes(120), SUMES0(120)
Dim ss, SEIRES_selidas As Integer
Dim MEGISTO_PLATOS As Integer

Dim metaf As String



MEGISTO_PLATOS = 120
For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next 'SEIR_SELID1

SEIRES_selidas = Val(FindParametroi("EKTYPOTES", "SEIR_SELID1", "70", "Σειρές ανά σελίδα(κάθετη)"))
' f_psifiaAjias = Val(FindParametroi("PAR1", "f_psifiaAjias", "2", "Δεκαδικά Ψηφία Αξίας Σειρών τιμολογιου")) 'posa psifia tha exei h kathe seira

Dim DUM, n
Dim F_T As New ADODB.Recordset

Dim TT As Long

TT = GetCurrentTime()


F_T.Open sql, gdb, adOpenForwardOnly, adLockReadOnly




n = F_T.Fields.Count



For k = 1 To n
    If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then 'DATE
      f_tab(k) = 2 + f_tab(k - 1) + 8
    Else
      If F_T.Fields(k - 1).DefinedSize > 200 Then
         f_tab(k) = 2 + f_tab(k - 1) + 13
      Else
         f_tab(k) = 3 + f_tab(k - 1) + F_T.Fields(k - 1).DefinedSize
      End If
    End If
Next


LSYN = f_tab(k - 1)





mPal = "    "
mPAL22 = "     "
MOLIS_ALLAJE = 0

If LSYN > MEGISTO_PLATOS Then
   SEIRES_selidas = 45
   Printer.Orientation = vbPRORLandscape '  2   Documents are printed with the top at the wide side of the paper.
Else
  'TO PAIRNEI APO TO TABLE PARAMETROI
   Printer.Orientation = vbPRORPortrait  ' 1   Documents are printed with the top at the narrow side of the paper.
   
End If






'If IsNull(EPIK) Then
 '   EPIK = Format(Date, "dd/mm/yyyy")
'End If

'marxeio2 = "EKT" + xeirisths + ".TXT"
On Error Resume Next

'Close #1
'Open "c:\print" For Output As #1
'  LSYN = 1
'  For K = 1 To n
'       LSYN = LSYN + 2 + F_T(K - 1).DefinedSize  'Len(macro("p", aa)) + 1
'  Next
  
  
'  If LSYN > 128 Then
'      sthles = LSYN + 10
'  ElseIf LSYN > 80 Then
'      sthles = LSYN + 10 ' 136
'  End If

   m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )
   For k = 1 To n
           aa = LTrim(Str(k))
          
          m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + 7 ' Len(macro("p", aa)) + 1
   Next
 
 
 
 Printer.FontSize = 8
 Printer.FontName = "Courier New"
 Printer.Font.Charset = 161
 'Printer.FontBold = True
 
 
 
 
 
 
' PRINT2_Xepik
    
    
Dim R As New ADODB.Recordset
R.Open "SELECT *FROM MEM", gdb, adOpenDynamic, adLockOptimistic
'R("pelono") = pel_ono.Text
'R("pelepa") = pel_epa.Text
'R("peldie") = pel_die.Text
 

    
         Printer.Print R("pelono")
         Printer.Print R("pelepa")

         Printer.Print Now
            Printer.FontSize = 12
    Printer.FontBold = True
     Printer.Print Tab(9 / 12 * (LSYN / 2 - Len(EPIKEF) / 2)); EPIKEF
    Printer.FontSize = 9
    '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------
    
    For k = 0 To n - 1
        Printer.Print Tab(f_tab(k)); Left(F_T(k).Name, F_T(k).DefinedSize);
    Next

    'δευτερη σειρα επικεφαλίδας
    For k = 0 To n - 1
        Printer.Print Tab(f_tab(k)); Mid$(F_T(k).Name, F_T(k).DefinedSize + 1, Len(F_T(k).Name) - F_T(k).DefinedSize + 1);
    Next
    Printer.FontBold = False
    Printer.Print Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")

For I = 1 To n
  m_sumes(I) = 0
  SUMES0(I) = 0
Next
Dim LAST_TIMH
Dim synt1
Dim maxWidth As Long

synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
AYJ = 0
Typose = 0
Dim row
Dim synola_SELIDOS
synola_SELIDOS = True ' False

'--------------------------------------------------------
Do While Not F_T.EOF
    AYJ = AYJ + 1
      
      If end_page = 2 Then ' ΤΥΠΩΝΕΙ ΓΡΑΜΜΟΥΛΕΣ --------------------------------
            Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN, "-")
         end_page = 0
         MOLIS_ALLAJE = 1
      End If ' end_page=2

      
    If synola_SELIDOS Then
      If (AYJ - 1) Mod SEIRES_selidas = 0 And AYJ > 1 Then ' .AND. PROTH_FORA<>1 )
        'syn_epik
         Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN, "-")
         metaf = "Σε μεταφορά "
         GoSub printSUM
         
         Printer.NewPage
         
         'ξαναδίνω το orientation γιατί το χάνει μετά το endDoc
         If LSYN > MEGISTO_PLATOS Then
                 Printer.Orientation = vbPRORLandscape '  2   Documents are printed with the top at the wide side of the paper.
         Else
                 Printer.Orientation = vbPRORPortrait  ' 1   Documents are printed with the top at the narrow side of the paper.
         End If
         
         
         
         end_page = 0
         
         '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------
         
        
         Printer.Print
         Printer.Print
         Printer.Print
                   
         
         
         For k = 0 To n - 1
            Printer.Print Tab(f_tab(k)); Left(F_T(k).Name, F_T(k).DefinedSize);
         Next
        'δευτερη σειρα επικεφαλίδας
        For k = 0 To n - 1
            Printer.Print Tab(f_tab(k)); Mid$(F_T(k).Name, F_T(k).DefinedSize + 1, Len(F_T(k).Name) - F_T(k).DefinedSize + 1);
        Next
         Printer.Print Tab(m_sthl_ektyp(1)); String$(LSYN - 2, "-")
         
         metaf = "Εκ μεταφοράς "
         GoSub printSUM

      End If  'end_page = 1 Then
    End If
    
      
      
     '----------------- ΤΥΠΩΣΕ ΟΛΑ ΤΑ ΠΕΔΙΑ ---------------------
    For k = 0 To n - 1
    
      If k < n - 1 Then
            ' για να μην παταει στην επόμενη στήλη
             maxWidth = f_tab(k + 1) - f_tab(k) - 1
      End If

    
      If Left(F_T(k), 3) = "@@@" Then '  A/A
          Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(AYJ, "######"), Min(F_T(k).DefinedSize, maxWidth));
      ElseIf F_T(k).Type = 7 Or F_T(k).Type = 5 Then   'IsNumeric(f_t(K)) πραγματικο
          Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(F_T(k), "######,##0.00"), Min(maxWidth, 14)); '  F_T(k).DefinedSize);
      ElseIf F_T(k).Type = 8 Then   'DATE
          Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(F_T(k), "DD/MM/YYYY"), Min(maxWidth, 10));
      Else ' 10 STRING
      
           If IsNull(F_T(k)) Then
             Printer.Print Tab(f_tab(k));
           Else
             ' PRINTER.PRINT Tab(f_tab(K)); to928(F_T(K));
             If k < n - 1 Then
                ' για να μην παταει στην επόμενη στήλη
                Printer.Print Tab(f_tab(k)); Left(to928(F_T(k)), maxWidth);
             Else
                Printer.Print Tab(f_tab(k)); to928(F_T(k));
             End If
           End If

           If m_sthl_ektyp(k) > 2 Then
             If k = 1 Then Printer.Print
           End If  'm_sthl_ektyp(K) > 2
           
       End If
       ' soymes---------------------
       If Mid$(SUgm_str, k + 1, 1) = "1" Or Mid$(SUgm_str, k + 1, 1) = "2" Then
               If IsNull(F_T(k)) Then
    
               Else
                  m_sumes(k) = m_sumes(k) + F_T(k)
               End If
       End If ' mid$(sugm_str,k,1)
    Next
      
      Printer.Print
      If GROUPN > 0 Then
         LAST_TIMH = F_T(GROUPN - 1)
      End If
    F_T.MoveNext
    If Not F_T.EOF Then
       If GROUPN > 0 Then
           If LAST_TIMH <> F_T(GROUPN - 1) Then
             Printer.Print
           End If
       End If
    End If
    
    
    
Loop


'"
 Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")
'"
  ' PRINTER.PRINT Chr(13)  ' 6/12/2007
'"
'   aa = f_kodik("sum_seltxt")
   Dim PARAM '
   PARAM = IIf(aa = " ", "  ", aa)
   'pr_SUMselidas param
   'PRINT2_Xsumes "ΓΕΝΙΚΟ ΣΥΝΟΛΟ"
   
            metaf = "Σύνολα "
   
   GoSub printSUM
   Printer.Print
   Printer.Print Tab(m_sthl_ektyp(1) + 1); String$(LSYN - 2, "-")
   
'   PRINTER.PRINT Chr(12)
'"
Dim YPOSEL(10)
   For k = 1 To 4
     Printer.Print    '  1, 0; YPOSEL(k)
   Next
   'PRINTER.PRINT Chr(13)

If eject = "y" Then
   'PRINTER.PRINT Chr(12)
End If

' αποθηκεύω τις σούμες στα e1,e2,.. για να τις χρησιμοποιώ
'  For K = 1 To N
 '   aa = "e" + LTrim(Str(K))
'    macro(aa, 0) = m_sumes(K)
 ' Next

'Close #1
   
   
   
   pp = 1
'   If pp = 1 Then
'     If ektypoths = -1 Then
'     Else
'
'
'       PPD = MsgBox("ΤΟ ΒΛΕΠΩ ΠΡΙΝ ΤΥΠΩΘΕΙ", vbYesNo)
'       If PPD = vbYes Then
'
'          On Error GoTo 0
'           DUM = Shell("c:\mercvb\notepad.exe c:\print", vbMaximizedFocus)
'
'           PPF = MsgBox("Προχωρώ στην εκτύπωση", vbYesNo)
'           If PPF = vbYes Then
'               DUM = Shell("c:\mercvb\notepad.exe /p c:\print", vbMaximizedFocus)
'           End If
'       Else
'       End If
'     End If
'   End If
   SELIS = 1
Printer.EndDoc

Exit Function


printSUM:

Printer.Print Tab(0); Trim(metaf);
For k = 1 To n - 1
    
      If k < n - 1 Then
            ' για να μην παταει στην επόμενη στήλη
             maxWidth = f_tab(k + 1) - f_tab(k) - 1
      End If
        
   
    If m_sumes(k) > 0 Then
       If Mid$(SUgm_str, k + 1, 1) = "1" Then  'ΑΘΡΟΙΣΜΑ
          Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k), "######,##0.00"), Min(maxWidth, 14));   ' F_T(K).DefinedSize));
       Else  ' ΜΕΣΟΣ ΟΡΟΣ
          Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k) / AYJ, "######,##0.00"), Min(maxWidth, F_T(k).DefinedSize));
       End If
    Else
      '   Printer.Print Tab(f_tab(K)); ""; ' Right(Space(50), F_T(K).DefinedSize);
    End If
Next
    Printer.Print
Return

printEpik:

For k = 0 To n - 1
      If k < n - 1 Then
            ' για να μην παταει στην επόμενη στήλη
             maxWidth = f_tab(k + 1) - f_tab(k) - 1
      End If
    
    
    If m_sumes(k + 1) > 0 Then
         Printer.Print Tab(f_tab(k)); Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), Min(maxWidth, F_T(k).DefinedSize));
    Else
         Printer.Print Tab(f_tab(k)); Right(Space(50), Min(maxWidth, F_T(k).DefinedSize));
    End If
Next
    Printer.Print
Return







End Function


Function macro(string_, aa)
Dim fcheckonly As Boolean

  gvar = "" ' gia na midenizo tin makroentoli
 ' macro = EbExecuteLine(StrPtr(string_), 0&, 0&, Abs(fcheckonly)) = 0
  
'  παράδειγμα χρησιμοποίησης μακροεντολής
'  --------------------------------------
  
'   Private Sub Command1_Click()
'   Dim res As Boolean, var As Single
' αν χρησιμοποιούντα recordset πρεπει να είναι GLOBAL όπως GEID,GPEL,κ.λ.π.
'   res = ExecuteLine("var=2+3*(5+6):a$=var")
'
'   End Sub





End Function


'
'proc CHECK_FIELDS
'*********************
'priv aa, BB
'              aa = StrZero(K, 2)
'              BB=trim(kodik("f&aa)
'
'            IF TYPE(BB)="U"
'                 wait 'προβλημα στο πεδίο '+  aa
'            Else
'              @ 24,22 say &BB
'              IF TYPE(BB)<>"N" .AND. synola(k)='1'
'                 synola(k)='0'
'              End If
'            End If
'Return











'
'#include "\clipper\include\error.ch"
'#include "\clipper\include\Fileio.ch"
'#include "\clipper\include\inkey.ch"
'
Function ektyp_part(m__rec, mFile)


'*****************************************************************************
'LOCAL bOldHandler,mRow,mm_Head1,mm_head2,mm_head3,mm_head4
'PARA m__rec,mFile    && 0 ta dino ola edo
'
'PRIV i,Val_Field(150),arr1(50),a,zp,sfalma:=0,apo,eos,sapo,sevs,sure,mTypos(25),mString(25)
'PRIV EPIKEF_OU,SELIS:=0,synola(30),mBSEIRA := space(25),mdiax:=0
'priv YPOSEL(10),eject:='y',f(30),dok
'priv m_ord(10),_a,files:=1,b,arr2(30),head(30),mP(30)
'priv m1,m2,m3,m0,m4,_mTot(6),_kt,_St
'
'
'
'  For K = 1 To 4
'     yposel(k)=''
'  Next
'
'
'  For K = 1 To 30
'     f (K) = 0
'  Next
'
'  set colo to  &wn,&nw,,,gr+/b  &&  w/bg+
'  Clear
'   _a=1
'
'do while .t.
'   Clear
'
'   sure='Ο'
'   store date() to apo,evs
'   xvrio = Space(20)
'    kau = Space(20)
'   mseir = Space(20)
'   ms4 = Space(20)
'   ms5 = Space(20)
'
'   select 40
'     NET_USE('FONTS',.F.,0)
'     K = 0
'     Do While .not.EOF()
'        mTypos(++k)=FONTS("Typos
'        mString(k)  =FONTS("String
'        Skip
'     enddo
'
'   SELECT 40
'
'   if pcount()<2
'      mFile='kodik'
'   End If
'
'     if .not. file(mFile+'.DBF')
'         do CHECK_KODIK with mFile
'     End If
'
'      NET_USE(mfile ,.F.,0,'kodik')
'      IF LASTREC()=0
'         For K = 1 To 20
'            ADD_REC (0)
'         Next
'      End If
'
'      IF FCOUNT()<132
'         do CHECK_KODIK with mFile
'         NET_USE(mFile , .F. , 0 ,'kodik' )
'      End If
'
'   mm_Head1=kodik("m_Head1
'   mm_Head2=kodik("m_Head2
'   mm_Head3=kodik("m_Head3
'   mm_Head4=kodik("m_Head4
'
'   index on epikef2 to dokkod&xeirisths
'
'   go top
'   n = 0
'   Do While .not.EOF()
'      K = recn()
'      n = n + 1
'      arr1(n)= left(kodik("epikef2,55)+'.. '+trans(K,'99')
'      Skip
'   enddo
'
'
'
'  if m__rec = 0
'      _a=achoice ( 3,7,23,68,arr1,,,_a)
'  Else
'      _a=m__rec
'  End If
'
'
'  Clear
'
'   EPIKEF_OU = if( _a=0,'',arr1(_a) )
'   @ 0,20 say EPIKEF_OU
'
'
'
'
'
'
'   @ 4,0 say 'Nα διορθώσω το report' get sure
'   Read
'   if lastkey()=27
'       Close Data
'       Return
'   End If
'
'
'********* ----------------------------------------------------------------
'
'
'
'
'   go Val ( Right(arr1 ( if(_a<=0,1,_a) )  , 3 ) )
'
'
'   if sure $ 'nNνΝ'
'      if arx1='   '  && αν είναι νέο record να ρωτά από πού θα αντιγραφεί
'           mReck = recn()
'           mR = 1
'           @ 6,0 say 'Να αντιγραφεί από το Report No ' get mR
'           Read
'        * διαβάζω το παλιό record
'        go mR
'        For i = 1 To fcount()
'            f = Field(i)
'            Val_Field(i) = &f
'        Next
'
'        go mReck
'        rec_lock (0)
'        * αντιγράφω στο καινούριο record
'        For i = 1 To fcount()
'            f = Field(i)
'            repl &f with Val_Field(i)
'        Next
'        unlock
'
'      End If
'      rec_lock (0)
'      Clear
'      @ 0,20 say EPIKEF_OU
'      set message to 24
'      * @ 2,0 say 'αριθμός αρχείων που θα ανοίξουν ' get Number_arx
'      @ 3,0 say 'Αρχείο κύριο(με / κατάλογο)' get arx1 pict '@S10'   valid find_arxeio(arx1)=0
'
'      @ 4,0 say 'Αρχείο 2ο  ' get arx2
'      @ 5,0 say 'Σχέση που τα συνδέει με το κύριο  ' get relation2
'
'      @ 7,0 say 'Αρχείο 3ο  ' get arx3
'      @ 8,0 say 'Σχέση που τα συνδέει με το κύριο  ' get relation3
'
'      @10,0 say 'Αρχείο 4ο  ' get arx4
'      @11,0 say 'Σχέση που τα συνδέει με το κύριο  ' get relation4
'
'
'      @15,0 say 'Συνθήκη άθροισης on..  ' get total pict '@S30'
'      Read
'
'
'      if total<>'   '
'         For K = 1 To 6
'             _mTot(k) = subs( pedia_tot , 01+(k-1)*10 , 10 )
'         Next
'
'         @ 15,0 say 'Πεδία άθροισης '
'         For K = 1 To 6
'             @ 16,(k-1)*10+k-1  get _mTot(k)
'         Next
'
'         @ 17,0 say 'συνθήκη άθροισης for..' get for_tot
'         Read
'
'         m_tot2=''
'         For K = 1 To 6
'             m_tot2 = m_tot2 + _mTot(k)
'         Next
'         repl pedia_tot with m_tot2
'
'      End If
'
'      repl Number_arx with if(arx2=' ',1,if(arx3=' ',2,if(arx4=' ',3,4)))
'      unlock
'
'
'
'
'
'
'   End If
'
'
'      mNumber_arx=kodik("Number_arx
'      mARX1=trim(kodik("arx1)
'      mARX2=trim(kodik("arx2)
'      mARX3=trim(kodik("arx3)
'      mARX4=trim(kodik("arx4)
'      mRelation4 = Trim(relation4)
'      mRelation2 = Trim(relation2)
'      mRelation3 = Trim(relation3)
'      mBSEIRA = KODIK("bSEIRA
'
'
'
'      do OPEN_DATA WITH .F.,1,mARX1
'      For K = 1 To 4
'          m_ord (K) = indexkey(K)
'      Next
'
'
'      *        net_use(mArx1,.f.,0)
'   if mNumber_arx-1 > 0
'      do OPEN_DATA WITH .F.,mNumber_arx-1,mARX2,mARX3,mARX4
'   End If
'
'    sele &mARX1
'    make_relations()
'
'   if sure $ 'nNνΝ'
'        Clear
'        @ 0,20 say EPIKEF_OU
'
'
'     *do while .t.
'        sele kodik
'        rec_lock (0)
'
'        if kodik("m_head1=space(20) .and. kodik("m_head2=space(20)
'           repl m_Head1 with mm_Head1
'           repl m_Head2 with mm_Head2
'           repl m_Head3 with mm_Head3
'           repl m_Head4 with mm_Head4
'        End If
'
'
'
'        @ 0,0 say 'Τίτλος Εκτύπωσης(οθόνη-μενού)' get kodik("epikef2 PICT '@S30'
'
'
'
'
'
'        mRow = 2
'        @ mRow+0,0 say 'α Τίτλος  (επάνω-αριστ)'  get kodik("m_Head1  PICT '@S30'
'        @ mRow+1,0 say 'β Τίτλος  (επάνω-αριστ)'  get kodik("m_Head2  PICT '@S30'
'        @ mRow+2,0 say 'γ Τίτλος  (επάνω-αριστ)'  get kodik("m_Head3  PICT '@S30'
'        @ mRow+3,0 say 'δ Τίτλος  (επάνω-αριστ)'  get kodik("m_Head4  PICT '@S30'
'
'
'
'
'        mRow = 6
'        @ mRow+0,0 say 'α Επικεφαλίδα Εκτύπωσης(κέντρο)'  get kodik("epikef  PICT '@S30'
'        @ mRow+1,0 say 'β Επικεφαλίδα Εκτύπωσης(κέντρο)'  get kodik("epik_2   PICT '@S30'
'        @ mRow+2,0 say 'γ Επικεφαλίδα Εκτύπωσης(κέντρο)'  get kodik("epik_3  PICT '@S30'
'
'
'
'        @ mRow+4,0 say 'Aριθμός πεδίων' get kodik("nr range 1,25
'
'        For K = 1 To 4
'          @ 15+k,0 say str(k)+'='+m_ord(k)
'        Next
'
'        @ mRow+5,35 say 'ταξινόμηση κατά '    get kodik("tag PICT '@S20'
'
'        @ mRow+6,0 say 'Συνθήκη << όσο συμβαινει τύπωνε >>  ' get kodik("syntwh
'        @ mRow+7,0 say 'Φίλτρο ' get kodik("syntIF PICT '@S70'
'        Read
'
'
'
'
'
'
'        @ 15,0 clear to 24,79
'
'
'        a=1  && βρισκω την προηγούμενη ρύθμιση των fonts
'        For K = 1 To Len(mString)
'            if mString(k)=trim ( KODIK("FONTS )
'               a = K
'               exit
'            End If
'        Next
'
'        @ 8,0 say 'Δώσε τον τύπο της εκτύπωσης'
'        * a=achoice ( 8,30,12,58,mTypos,,,a )
'        repl kodik("fonts with if( a=0,'',mString(a))
'
'
'        if type(kodik("syntwh)='UI' .or. type(kodik("syntwh)='UE'
'           wait 'προβλημα στην συνθήκη << όσο συμβαινει τύπωνε >> '
'           BREAK
'        End If
'        if type(kodik("syntIF)='UI' .or. type(kodik("syntIF)='UE'
'           wait 'προβλημα στο φίλτρο '
'           BREAK
'        End If
'
'        Clear
'        @ 0,20 say EPIKEF_OU
'
'        @ 1, 0  say 'Επικ.Πεδίου'
'        @ 1,22 say 'Τίτλος Πεδίου'
'        @ 1,44 say 'format γραφής'
'        @ 1,70 say 'Σύνολο=1'
'
'      SELECT 41
'         NET_USE("FIELDS",.F.,0)  && ΟΝΟΜΑΤΑ ΠΕΔΙΩΝ
'            set filter to trim(arxeio) $ mARX1+mARX2+mARX3+mARX4
'
'
'        set key 14 to insert_line  && Ctrl + N
'        set key 25 to delete_line  && Ctrl + Y
'        for k=1 to kodik("nr
'            synola(k)=subs(kodik("synolo,k,1)
'        Next
'        set key -9 to opos_allo
'        for k=1 to min(22,kodik("nr)
'            aa = StrZero(K, 2)
'            @ 1+k, 0 say k pict '99'
'            @ 1+k, 3 get kodik("ep&aa pict '@S20' valid find_field(kodik("ep&aa,MARX1,MARX2,MARX3)=0
'            @ 1+k,25 get kodik("f&aa  pict '@S20'
'            @ 1+k,47 get kodik("p&aa  pict '@S20'
'            @ 1+k,70 get synola(k)    pict '9'
'        Next
'        Read
'        set key 14 to  && Ctrl + N
'        set key 25 to  && Ctrl + Y
'
'
'        Clear
'        if kodik("nr > 22
'          for k=23 to kodik("nr
'            aa = StrZero(K, 2)
'            @ k-21, 0 say k pict '99'
'            @ k-21, 3 get kodik("ep&aa pict '@S20' valid find_field(kodik("ep&aa)=0
'            @ k-21,25 get kodik("f&aa  pict '@S20'
'            @ k-21,47 get kodik("p&aa  pict '@S20'
'            @ k-21,70 get synola(k)    pict '9'
'          Next
'          Read
'        End If
'
'        set key -9 to
'
'        sele &mARX1
'        boldhandler:=errorblock({|e| MyERRORHANDLER(e,boldHandler)})
'        begin sequence
'          * σε περίπτωση που χρησιμοποιώ μία άλλη στήλη π.χ. f6 να μην βγάζει λάθος
'          ssss = 1
'          for k=1 to kodik("nr
'             aa = alltrim(Str(K))
'             f&aa='ssss'
'          Next
'          for k=1 to kodik("nr
'            DO CHECK_FIELDS
'          Next
'        recover
'            aa = StrZero(K, 2)
'            IF TYPE("KODIK("F"+AA)="U"
'                 wait 'προβλημα στο πεδίο '+  aa
'            Else
'                 aa=trim(kodik("f&aa)
'                 wait 'προβλημα στο πεδίο '+ STR(K) + '  '+ aa
'            End If
'          ERRORBLOCK (boldhandler)
'        end sequence
'        ERRORBLOCK (boldhandler)
'
'        @ 22,0 say '2η σειρά:' get kodik("bSEIRA PICT '@S40'
'        @ 23,0 say 'Aλλάζει σελίδα και όταν αλλάζει το string πεδίο Νο π.χ. 2  0=δεν αλλάζει' get kodik("s_eject
'        @ 24,0 say 'Θά αλλάζει σελίδα=0 διαχωρισμός με παύλες=1 ' get kodik("diax
'        Read
'        mBSEIRA = KODIK("bSEIRA
'        Clear
'
'        if kodik("s_eject > 0
'            IF kodik("synt_alag='  '
'               REPL  kodik("synt_alag WITH 'mpal=mpal2'
'            End If
'            @ 2,0 say 'Συνθήκη για αλλαγή σελίδας (mpal=mpal2)' get kodik("synt_alag pict '@S20'
'            @ 4,0 say 'Κείμενο που θά γράφεται στά αθροίσματα' get kodik("sum_seltxt
'        End If
'        @ 6, 0 say 'Νά τυπώνει αθροίσματα σε μεταφορά=1   οχι=0 ' get kodik("se_metaf
'        @ 8, 0 say 'Νά τυπώνει αθροίσματα εκ μεταφοράς=1  όχι=0 ' get kodik("ek_metaf
'        Read
'
'        set key 14 to
'        set key 25 to
'        mSYNOLO=''
'        for k=1 to kodik("nr
'            mSYNOLO=mSYNOLO+synola(k)
'        Next
'        REPL kodik("SYNOLO WITH MSYNOLO
'        Clear
'        @ 2,0 say 'Κείμενο στην 1η ημερομηνία(apo)  ' get  kodik("k_hm1
'        @ 3,0 say 'Κείμενο στην 2η ημερομηνία(evs)  ' get  kodik("k_hm2
'        @ 5,0 say 'Κείμενο στo  1o όνομα     (xvrio)' get  kodik("k_ch1
'        @ 6,0 say 'Κείμενο στo  2ο όνομα     (kau)  ' get  kodik("k_ch2
'        @ 7,0 say 'Κείμενο στo  3ο όνομα     (mseir)' get  kodik("k_ch3
'        @ 8,0 say 'Κείμενο στo  4ο όνομα     (ms4)  ' get  kodik("k_ch4
'        @ 9,0 say 'Κείμενο στo  5ο όνομα     (ms5)  ' get  kodik("k_ch5
'        Read
'*        unlock
'
'
'        *EXIT
'     *enddo
'
'
'   End If
'
'
'  Clear
'  if kodik("k_hm1<>'  '
'     @ 1, 0 say trim(kodik("k_hm1) get apo
'  End If
'  if kodik("k_hm2<>'  '
'     @ 2, 0 say trim(kodik("k_hm2) get evs
'  End If
'  if kodik("k_ch1<>'  '
'     @ 4, 0 say trim(kodik("k_ch1) get xvrio
'  End If
'
'  if kodik("k_ch2<>'  '
'     @ 5, 0 say trim(kodik("k_ch2) get kau
'  End If
'
'  if kodik("k_ch3<>'  '
'     @ 6, 0 say trim(kodik("k_ch3) get mseir
'  End If
'
'  if kodik("k_ch4<>'  '
'     @ 7, 0 say trim(kodik("k_ch4) get ms4
'  End If
'
'  if kodik("k_ch5<>'  '
'     @ 8, 0 say trim(kodik("k_ch5) get ms5
'  End If
'
'  Read
'
'     if lastkey()=27
'        Loop
'     End If
'
'     Sapo = dtoc(apo)
'     Sevs = dtoc(evs)
'
'     synt_eject=kodik("s_eject
'
'
'   sele &mARX1
'
'
'    Err = 0
'
'    if err = 1
'       Close Data
'       Loop
'    End If
'
'
'
'    ZP = 1
'    do epilogh2 with 23,zp,3,'εκτύπωση σε οθόνη','εκτύπωση σε εκτυπωτή','edit'
'    IF LASTKEY()=27
'       Close Data
'       Loop
'    End If
'
'       b=trim(kodik("tag)
'
'
'       sele &marx1
'       if mARX1='MIS'
'          do make_MIS_totals
'       Else
'          do make_totals
'       End If
'
'
'
'
'
'
'
'
'
'
'       go top
'
'    progr=if(zp=1,"printbox4","print3_xar")
'
'
'*/*
'    IF ZP=3
'        xxx=kodik("syntIF
'        set filter to &xxx
'        go top
'        arr2 (1) = "lines"
'        for k=1 to kodik("nr
'            aa = StrZero(K, 2)
'            arr2(k)=kodik("f&aa
'            head(k)=kodik("ep&aa
'            mP(k)  =kodik("p&aa
'        Next
'        dbedit ( 1,0,24,79,arr2,"db17_udf",,head )
'        Close Data
'        Return
'    End If
'*/
'
'
'    IF ZP=2
'       @ 24,0 CLEAR TO 24,79
'       @ 24,0 SAY 'ΑΠΟ ΣΕΛΙΔΑ ' GET SELIS
'       Read
'    End If
'
'Clear
'  sthles = 200
'  epik=trim(kodik("epikef)
'  if len(trim(kodik("epikef))<>0
'     epik=&epik
'  End If
'  print_fonts=kodik("fonts
'
'  epik2 = trim(kodik("epik_2)
'  if len(trim(epik2))<>0
'      epik2 = &epik2
'  End If
'
'  epik3 = trim(kodik("epik_3)
'  if len(trim(epik3))<>0
'      epik3 = &epik3
'  End If
'
'  do &PROGR  with KODIK("NR,TRIM(KODIK("SYNTWH),TRIM(KODIK("SYNTIF),TRIM(KODIK("SYNOLO),;
'       TRIM(kodik("EP01),TRIM(kodik("F01),TRIM(kodik("P01),;
'       TRIM(kodik("EP02),TRIM(kodik("F02),TRIM(kodik("P02),;
'       TRIM(kodik("EP03),TRIM(kodik("F03),TRIM(kodik("P03),;
'       TRIM(kodik("EP04),TRIM(kodik("F04),TRIM(kodik("P04),;
'       TRIM(kodik("EP05),TRIM(kodik("F05),TRIM(kodik("P05),;
'       TRIM(kodik("EP06),TRIM(kodik("F06),TRIM(kodik("P06),;
'       TRIM(kodik("EP07),TRIM(kodik("F07),TRIM(kodik("P07),;
'       TRIM(kodik("EP08),TRIM(kodik("F08),TRIM(kodik("P08),;
'       TRIM(kodik("EP09),TRIM(kodik("F09),TRIM(kodik("P09),;
'       TRIM(kodik("EP10),TRIM(kodik("F10),TRIM(kodik("P10),;
'       TRIM(kodik("EP11),TRIM(kodik("F11),TRIM(kodik("P11),;
'       TRIM(kodik("EP12),TRIM(kodik("F12),TRIM(kodik("P12),;
'       TRIM(kodik("EP13),TRIM(kodik("F13),TRIM(kodik("P13),;
'       TRIM(kodik("EP14),TRIM(kodik("F14),TRIM(kodik("P14),;
'       TRIM(kodik("EP15),TRIM(kodik("F15),TRIM(kodik("P15),;
'       TRIM(kodik("EP16),TRIM(kodik("F16),TRIM(kodik("P16),;
'       TRIM(kodik("EP17),TRIM(kodik("F17),TRIM(kodik("P17),;
'       TRIM(kodik("EP18),TRIM(kodik("F18),TRIM(kodik("P18),;
'       TRIM(kodik("EP19),TRIM(kodik("F19),TRIM(kodik("P19),;
'       TRIM(kodik("EP20),TRIM(kodik("F20),TRIM(kodik("P20),;
'       TRIM(kodik("EP21),TRIM(kodik("F21),TRIM(kodik("P21),;
'       TRIM(kodik("EP22),TRIM(kodik("F22),TRIM(kodik("P22),;
'       TRIM(kodik("EP23),TRIM(kodik("F23),TRIM(kodik("P23),;
'       TRIM(kodik("EP24),TRIM(kodik("F24),TRIM(kodik("P24),;
'       TRIM(kodik("EP25),TRIM(kodik("F25),TRIM(kodik("P25),;
'       TRIM(kodik("EP26),TRIM(kodik("F26),TRIM(kodik("P26),;
'       TRIM(kodik("EP27),TRIM(kodik("F27),TRIM(kodik("P27),;
'       TRIM(kodik("EP28),TRIM(kodik("F28),TRIM(kodik("P28),;
'       TRIM(kodik("EP29),TRIM(kodik("F29),TRIM(kodik("P29),;
'       TRIM(kodik("EP30),TRIM(kodik("F30),TRIM(kodik("P30)
'Close Data
'       SELIS = 1
'
'dok='dok'+xeirisths+'.mdx'
'erase &dok
'
'
'
'enddo
'
'
'Return
End Function
'
'
'Function find_arxeio(dum)
'*********************************************
'priv arxeia(9),Marxeia(9),k,l,sc
'
'if dum<>'/'
'   return 0
'End If
'
'  sc = savescreen(4, 40, 7, 79)
'      arxeia(1)='ERGAZ'
'
'    /*
'       arxeia(2)='TIM'
'       arxeia(3)='GRA'
'       arxeia(4)='EID'
'       arxeia(5)='EGG'
'       arxeia(6)='EGGTIM'
'    */
'
'      Marxeia(1)='Αρχείο Εργαζόμενου'
'    /*
'      Marxeia(2)='Αρχείο Παραστατικών'
'      Marxeia(3)='Αρχείο Επιταγών / Γραμματίων'
'      Marxeia(4)='Αρχείο ειδών '
'      Marxeia(5)='Αρχείο κινήσεων πελατών/προμηθευτών '
'      Marxeia(6)='Αρχείο κινήσεων ειδών'
'     */
'
'For K = 1 To 6
'    @ k+3,40 prompt arxeia(k) message Marxeia(k)
'Next
'menu to l
'restscreen (4,40,7,79,sc)
' if l<>0
'    repl arx1 with arxeia(l)
' End If
'return 0
'
'
'
'
'
'proc PRINTBOX4
'*********************** P R I N T B O X ************************
'**************      κάνει παρουσίαση αρχείου
'*** n=αριθμός fields που παρουσιάζονται,synuhkh συνθήκη για το DO WHILE
'*** synuhkh2 η συνθήκη για το  IF, sugm_str βλέπει που θα κάνει σούμες ,π.χ. "00100" κάνει σούμες στο 3ο field
'*** sum_pic  το picture γιά τις σούμες
'*** Ei,Fi,Pi  :Επικεφαλίδα παρουσίασης,Fields που παρουσιάζονται,Picture παρουσίασης
'
'para n,synuhkh,synuhkh2,sugm_str, _e1,f1,p1, _e2,f2,p2, _e3,f3,p3, _e4,f4,p4, _e5,f5,p5,  _e6,f6,p6, _e7,f7,p7, _e8,f8,p8, _e9,f9,p9 ,_e10,f10,p10,_e11,f11,p11, _e12,f12,p12,_e13,f13,p13,_e14,f14,p14,_e15,f15,p15,_e16,f16,p16, _e17,f17,p17,_e18,f18,p18,_e19,f19,p19,_e20,f20,p20,;
'          _e21,f21,p21, _e22,f22,p22, _e23,f23,p23, _e24,f24,p24, _e25,f25,p25,_e26,f26,p26, _e27,f27,p27,_e28,f28,p28,_e29,f29,p29,_e30,f30,p30
'priv sumes, synuhkh, synuhkh2, sugm_str, cc, pp, AYJ, P
'
'
'declare S(30),sumes(30),OUON(5)
'For K = 1 To 30
'   sumes (K) = 0
'Next
'For K = 1 To 5
'   ouon(k)=''
'Next
'   LSYN = 1
'   epikef0 = "Ι"
'   epikef = "Ί"
'   epikef2 = "Μ"
'   patos = "Θ"
'   GEMIS = "Ί"
'For K = 1 To n
'   aa = alltrim(Str(K))
'   lsyn = lsyn + len(p&aa) + 1
'   epikef0 = epikef0 + repl("Ν",len(p&aa))+"Λ"
'   epikef  =epikef +left(_e&aa+space(50),len(p&aa))+"Ί"
'   epikef2 = epikef2 + repl("Ν",len(p&aa))+"Ξ"
'   PATOS = patos + repl("Ν",len(p&aa))+"Κ"
'   GEMIS = gemis + left(space(50),len(p&aa)) + "Ί"
'Next
'
' do printbox21
'Return
'
'
'
'static proc printbox21
'****************** συνέχεια απο παραπάνω ***********************************
'priv aa,pict(30),f(30)
'synt1 = if ( type("SYNT1")="U",'.t.',synt1)   && όταν έρχεται απο τnν αποθήκη ορίζεται το synt1
'
'm_SEIRA = 1
'
'
'
'
'Clear
'
'arxh0 = 1 + Len(p1 + p2) + 2
'arxh = arxh0 + 1
'nouon = 0
'
'ARXSTHL = IF ( LSYN>77 , 0 , 40 - LSYN/2 )
'
'     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
'     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
'     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)
'
'AYJ = 0
'do while &synuhkh
'     set cursor off
'
'
' IF &synuhkh2   &&  η δεύτερη συνθήκη / αν δεν υπάρχει τότε .Τ.
'
'   if &synt1   && η συνθήκη της αποθήκης αλλιώς .Τ.
'      AYJ = AYJ + 1
'
'
'          For K = 1 To n
'            aa = 'f'+ alltrim( str ( k ) )
'            aa = &aa
'            f(k)= &aa
'            aa = 'p'+ alltrim( str ( k ) )
'            aa = &aa
'            pict (K) = aa
'
'            ww = alltrim(Str(K))
'            aas = f&ww
'            sumes(k) = sumes(k) +  if ( subs(sugm_str,k,1)="1" , &aas  , 0 )
'          Next
'
'          u = "Ί"
'          For K = 1 To n
'             kena = len(pict(k))  - len( trans(f(k),pict(k))  )
'             kena = if ( kena=0 ,'' , space(kena) )
'             u = u + trans(f(k),pict(k))+kena+"Ί"
'          Next
'          s (m_SEIRA) = u
'
'      @ m_SEIRA+2,ARXSTHL SAY left(s(m_SEIRA),arxh0) + subs ( S(m_SEIRA) , arxh,79-arxh0 )
'      m_SEIRA = m_SEIRA + 1
'
'
'    endif  && synt1
' ENDIF  &&  η δεύτερη συνθήκη synuhkh2
'
'if type('TELOS') <> 'U'  &&   αν ορίστηκε το telos
'      if recn()=telos  && αν έφτασε στην τελευταία εγγραφή , να το κάνει eof()
'           seek 'ωωωωωωωωωωωωωωω'
'      Else
'           Skip
'      End If
'Else
'      Skip
'End If
'
'if ROW() >= 22
' @ ROW()+1,ARXSTHL SAY left(patos,arxh0)+SUBS(PATOS,ARXH,79-arxh0)
' set cursor on
' sumarisma = ''
' For K = 1 To n
'        ww = alltrim(Str(K))
'        aas = p&ww
'        sumarisma = sumarisma + " " +  if ( subs(sugm_str,k,1)="1" , trans(sumes(k),"&aas"),space( len(aas) )  )
' Next
' sumarisma = sumarisma + ' '
' @ 24,ARXSTHL say left(sumarisma,arxh0)+subs ( sumarisma , arxh , 79-arxh0 )
'
' nouon = nouon + 1
' if nouon > 5
'    For K = 1 To 4
'        ouon(k) = ouon(k+1)
'    Next
'    save screen to ouon(5)
'    nouon = 5
' Else
'    save screen to ouon(nouon)
' End If
'
' @ 24,ARXSTHL say "< Esc > επιστρέφω"
' DO Nwait
' plhk = lastkey()
' piso = 0
' do while (plhk=4 .or. plhk=19  .or. plhk=18 .or. plhk=3) .AND. lsyn>80 && LSYN-n>79 && (" ή <- ή PgDn ή PgUp
'   if plhk=18 && PgUp
'      piso = Min(piso + 1, nouon - 1)
'      restore screen from ouon(nouon-piso)
'   End If
'   if plhk=3 && PgDn
'      piso = Max(piso - 1, 0)
'      restore screen from ouon(nouon-piso)
'   End If
'   if plhk=4 && (" belos
'     arxh = Min(arxh + 2, LSYN - 78 + arxh0)
'     DO HOR_SCROLL
'   End If
'   if plhk=19 && <- belos
'     arxh = Max(arxh - 2, arxh0 + 1)
'     DO HOR_SCROLL
'   End If
'   do nwait
'   plhk = lastkey()
' enddo
'   Clear
'   m_SEIRA = 1
'   if lastkey()=27
'      Return
'   End If
'endif      &&  row() = 22
'
'     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
'     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
'     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)
'
'
'
'enddo
'
' Do While Row() < 22
'     s (m_SEIRA) = GEMIS
'     @ m_SEIRA+2,ARXSTHL SAY left(s(m_SEIRA),arxh0) + subs ( S(m_SEIRA) , arxh,79-arxh0 )
'     m_SEIRA = m_SEIRA + 1
' enddo
'
' @ ROW()+1,ARXSTHL SAY left(patos,arxh0)+SUBS(PATOS,ARXH,79-arxh0)
' set cursor on
' sumarisma = ''
' For K = 1 To n
'        ww = alltrim(Str(K))
'        aas = p&ww
'        sumarisma = sumarisma + " " +  if ( subs(sugm_str,k,1)="1" , trans(sumes(k),"&aas"),space( len(aas) )  )
' Next
' sumarisma = sumarisma + ' '
' @ 24,ARXSTHL say left(sumarisma,arxh0)+subs ( sumarisma , arxh , 79-arxh0 )
'
' nouon = nouon + 1
' if nouon > 5
'    For K = 1 To 4
'        ouon(k) = ouon(k+1)
'    Next
'    save screen to ouon(5)
'    nouon = 5
' Else
'    save screen to ouon(nouon)
' End If
'
' @ 24,ARXSTHL say "< Esc > επιστρέφω"
' DO Nwait
' plhk = lastkey()
' piso = 0
' do while plhk=4 .or. plhk=19  .or. plhk=18 .or. plhk=3 && (" ή <- ή PgDn ή PgUp
'   if plhk=18 && PgUp
'      piso = Min(piso + 1, nouon - 1)
'      restore screen from ouon(nouon-piso)
'   End If
'   if plhk=3 && PgDn
'      piso = Max(piso - 1, 0)
'      restore screen from ouon(nouon-piso)
'   End If
'   if plhk=4 && (" belos
'     arxh = Min(arxh + 2, LSYN - 78 + arxh0)
'     DO HOR_SCROLL
'   End If
'   if plhk=19 && <- belos
'     arxh = Max(arxh - 2, arxh0 + 1)
'     DO HOR_SCROLL
'   End If
'   do nwait
'   plhk = lastkey()
' enddo
'   Clear
'   m_SEIRA = 1
'   if lastkey()=27
'      Return
'   End If
'Return
'
'
'
'
'
'
'
'
'static PROC HOR_SCROLL
'****************** ΟΡΙΖΟΝΤΙΟ SCROLLING *************************
'SET CURSOR OFF
'priv K
'     @ 0,ARXSTHL SAY left(EPIKEF0,ARXH0) + SUBS(EPIKEF0,ARXH,79-arxh0)
'     @ 1,ARXSTHL SAY left(EPIKEF ,ARXH0) + SUBS(EPIKEF ,ARXH,79-arxh0)
'     @ 2,ARXSTHL SAY left(EPIKEF2,ARXH0) + SUBS(EPIKEF2,ARXH,79-arxh0)
'     For K = 3 To 22
'        @ k,ARXSTHL SAY left(s(k-2),arxh0) + subs ( S(k-2) , arxh,79-arxh0 )
'     Next
'
'     @ 23,ARXSTHL SAY left(patos,arxh0) + SUBS(PATOS,ARXH,79-arxh0)
'     @ 24,ARXSTHL say left(sumarisma , arxh0) + subs ( sumarisma , arxh , 79-arxh0 )
'SET CURSOR ON
'Return
'
'
'static proc nwait
'**********  waiting untill key pressed *************
'a = inkey()
'Do While a = 0
'  *@ 23,78 SAY ""
'  a = inkey()
'enddo
'Return
'
'
Sub syn_epik()


'********** εκτυπωση επικεφαλίδων & συνόλων *************************
'priv M
'     m=kodik("sum_seltxt
'     printer.print+2,1 say ''
'
'       do pr_SUMselidas with if(len(trim(m))=0,' ',&m)
'     if kodik("se_metaf=1
'        do print2_xSUMES  with 'ΣΥΝΟΛΑ ΣΕ ΜΕΤΑΦΟΡΑ'
'     End If
'   ** να μηδενίζει τους αθροιστές μόνο όταν αλλάζει το εν λόγω πεδίο
'   if synt_eject > 0
'      if end_Page = 1
'         for i=1 to n  && αποθηκεύω τα σύνολα για να μπορώ να βγάλω μετά τα συνολα σελίδας
'           sumes0(i) = sumes(i)
'         Next
'      End If
'   Else
'        for i=1 to n  && αποθηκεύω τα σύνολα για να μπορώ να βγάλω μετά τα συνολα σελίδας
'           sumes0(i) = sumes(i)
'        Next
'   End If
'
'
'
'   For K = 1 To 4
'     printer.print+1,0 say YPOSEL(k)
'   Next
'
'   eject
'   For kw = 1 To 4
'       if len ( ar_Print(kw) ) > 0
'          printer.print+1 , 0 say ar_Print(kw)
'       End If
'   Next
'
' IF ! type("EPIK")="U"
'   printer.print+1,IF(type("STHLES")="U",40,LSYN/2) - LEN(EPIK)/2  SAY EPIK
'   SELIS = SELIS + 1
'   printer.print,lsyn say SELIS
'
'   if epik2='  '
'      printer.print+1,IF(type("STHLES")="U",40,LSYN/2) - LEN(EPIK)/2  SAY  repl('-',len(epik) )
'   Else
'      printer.print+1,IF(type("STHLES")="U",40,LSYN/2) - LEN(TRIM(EPIK2))/2  SAY EPIK2
'      if epik3 <> '  '
'         printer.print+1,IF(type("STHLES")="U",40,LSYN/2) - LEN(TRIM(EPIK3))/2  SAY  epik3
'      End If
'   End If
' End If
'
'   do print2_xEPIK
'   printer.print+1,_sthl_ektyp(1)+1 SAY repl('-',lsyn-2)
'   * printer.print+1,0 SAY ''
'   ?? CHR(13)
'   if kodik("ek_metaf=1
'      do print2_xSUMES  with 'ΣΥΝΟΛΑ ΑΠΟ ΜΕΤΑΦΟΡΑ'
'   End If
End Sub
'
'
'
'
'
'
Sub pr_SUMselidas(TITLOS)

'   ******************************** τυπώνει τις σούμες SELIDOS για το print_xar ********
'   para TITLOS
'   printer.print+1,1+_sthl_ektyp(1) SAY TITLOS
'      For K = 1 To n
'       aa = alltrim(Str(K))
'           aaP='P'+aa
'           aaP2=&aaP
'       if subs(sugm_str,k,1)="1"              && τυπώνει τις σούμες
'            printer.print,_sthl_ektyp(k)+1 say sumes(k)-SUMES0(k) pict "&aap2"
'       End If
'      Next

End Sub

Sub PRINT2_Xsumes(txt)


'   ******************************** τυπώνει τις σούμες για το print_xar ********
'*  para txt
'   printer.print+1,1+_sthl_ektyp(1) SAY ''   &&  txt
'   For K = 1 To n
'       aa = alltrim(Str(K))
'           aaP='P'+aa
'           aaP2=&aaP
'
'
'       if subs(sugm_str,k,1)="1"              && τυπώνει τις σούμες
'           if _sthl_ektyp(k) > 2
'              printer.print+if(k=1,1,0) , _sthl_ektyp(k) say "³" && διαχωριστικο
'           End If
'
'
'          @  PROW(),_sthl_ektyp(k)+1 say sumes(k) pict "&aap2"
'       End If
'   Next
End Sub
'
'
'static proc OLDPRINT2_Xepik
'   ***************** τυπώνει επικεφαλίδες απο το print_xar
'  Local xor
'   printer.print+1,0 say ''
'   For K = 1 To n
'       aa = alltrim(Str(K))
'       xor=at(";",_E&aa)-1 && ψηφία από πρώτο κομμάτι που θα τυπωθεί
'       xor = if ( xor=0 , len(trim(_E&aa)) , xor ) && αν δεν έχει χώρισμα το παίρνει όλο
'       xor = if ( xor > len( P&aa ) , len( P&aa ) , xor )
'       @  prow(),_sthl_ektyp(k)+1 say  subs(_E&aa,1,xor)
'   Next
'   printer.print+1,0 say ''
'   For K = 1 To n
'       aa = alltrim(Str(K))
'       xor = at (";",_E&aa) && se ποιό σημείο υπάρχει το ;
'       mhk=len(trim(_E&aa))-xor && μήκος εκτύπωσης
'       mhk= if ( mhk > len( P&aa ) , len( P&aa ) , mhk )
'       @  prow(),_sthl_ektyp(k)+1 say  subs(_E&aa,xor+1,mhk)
'   Next
'Return
'
Sub PRINT2_Xepik()


'***************** τυπώνει επικεφαλίδες απο το print_xar   3 SEIRES EPIKEFALIDA
' Local xor,xor2, _EPI(40,3),k,seires_ep,l,arxh
' priv aa
'
'* ΧΩΡΙΖΩ ΤΗΝ ΕΠΙΚΕΦΑΛΙΔΑ ΣΕ 3 ΜΕΡΗ
'   seires_ep = 1
'   For K = 1 To n
'       aa = alltrim(Str(K))
'       xor=at(";",_E&aa) && το 1ο ;
'       if xor=0  && μια σειρά μόνο
'          _epi(k,1)=""
'          _epi(k,2)=""
'          _epi(k,3)=_E&aa
'       Else
'          xor2 = Rat(";",_E&aa) && το 2ο ;
'
'          if xor2 > xor   && εχω 2 ερωτηματικά ( 3 σειρές)
'              _epi(k,1)=subs( _E&aa , 1 , xor-1 )
'              _epi(k,2)=subs( _E&aa , xor+1  , xor2- xor -1 )
'              _epi(k,3)=subs( _E&aa , xor2+1 ,len(_E&aa) - xor2 -1 )
'              seires_ep = 3
'          else   && ολο και όλο 1  ; ( 2 σειρές )    xor2=xor1
'              _epi(k,1)=""
'              _epi(k,2)=subs( _E&aa , 1  , xor -1 )
'              _epi(k,3)=subs( _E&aa , xor+1 ,len(_E&aa) - xor -1 )
'              seires_ep = IIf(seires_ep = 1, 2, seires_ep)
'
'          End If
'       End If
'   Next
'
'
'
'
'arxh  = if(seires_ep=2,2,if(seires_ep=1,3,1))
'
'For l = arxh To 3
'
'    printer.print+1,0 say ''
'   For K = 1 To n
'
'       aa = alltrim(Str(K))
'
'           if _sthl_ektyp(k) > 2
'              printer.print+if(k=1,1,0) , _sthl_ektyp(k) say "³" && διαχωριστικο
'           End If
'
'           if len( _epi(k,l) )  > len( P&aa )
'              @  prow(),_sthl_ektyp(k)+1 say  Left ( _epi(k,l) , len( P&aa ) )
'           Else
'              @  prow(),_sthl_ektyp(k)+1 say  _epi(k,l)
'           End If
'
'   Next
'Next
'
End Sub
'
'
'
'Function Find_Field(dum, MARX1, MARX2, MARX3)
' ****************************************************
' LOCAL ppw,scr2
' declare arr1(50)
'
'   if left( dum , 1 ) $ '*/.'
'
'       save screen to scr2
'
'
'
'
'     IF upper(MARX1) $ 'ERGAZ-MIS'
'       ppw = 1
'       do epilogh3 with 18,40,ppw,2,'ΚΥΡΙΑ ΣΤΟΙΧΕΙΑ ΕΡΓΑΖΟΜΕΝΟΥ','ΣΤΟΙΧΕΙΑ ΜΙΣΘΟΔΟΣΙΑΣ'
'       if ppw=2
'          find_orisma()
'          restore screen from scr2
'          return 0
'       End If
'     End If
'
'       sele Fields
'       SET FILTER TO TRIM(ARXEIO) $ MARX1+MARX2+MARX3
'       GO TOP
'       For i = 1 To Min(fcount(), 50)
'         arr1 (i) = Field(i)
'       Next
'
'       dbedit ( 5,0,23,79,arr1)
'
'       restore screen from scr2
'
'       pa = Right(Trim(readvar()), 2)
'
'       sele kodik
'
'       repl ep&pa with Fields("epikef
'       repl f&pa with  trim(Fields("arxeio)+'("'+Fields("Field_Name
'       repl p&pa with  Fields("picture
'
'   End If
'
'return 0
'
'
'Function find_orisma()
'****************************************
'
'RETURN NIL
'
'
'
'
'
'
'proc insert_line
'*************************************************
'Local GetList:={}
'priv no_line:=0,k,aa,prev
'priv LAST_COLOR:=SETCOLOR()
'       no_line = Val(Right(Trim(readvar()), 2))
'       no_line = Max(1, no_line)
'       no_line = Min(15, no_line)
'       For K = 15 To no_line + 1 Step -1
'           aa = StrZero(K, 2)
'           prev = StrZero(K - 1, 2)
'           repl kodik("ep&aa with kodik("ep&prev
'           repl kodik("f&aa with kodik("f&prev
'           repl kodik("p&aa with kodik("p&prev
'       Next
'
'           aa = StrZero(no_line, 2)
'           repl kodik("ep&aa with ' '
'           repl kodik("f&aa with ' '
'           repl kodik("p&aa with '  '
'
'
'        ** ζωγραφίζω τα get
'        set color to &nw
'        for k=1 to kodik("nr
'            aa = StrZero(K, 2)
'            @ 1+k, 0 say k pict '99'
'            @ 1+k, 3 get kodik("ep&aa pict '@S20' valid find_field(kodik("ep&aa)=0
'            @ 1+k,25 get kodik("f&aa  pict '@S20'
'            @ 1+k,47 get kodik("p&aa  pict '@S20'
'            @ 1+k,70 get synola(k)    pict '9'
'        Next
'        SETCOLOR (LAST_COLOR)
'Return
'
'
'
'proc delete_line
'*************************************************
'Local GetList:={}
'priv no_line:=0,k,aa,epomeno
'priv LAST_COLOR:=SETCOLOR()
'       no_line = Val(Right(Trim(readvar()), 2))
'       no_line = Min(15, no_line)
'       no_line = Max(1, no_line)
'
'       For K = no_line To 22
'           aa = StrZero(K, 2)
'           epomeno = StrZero(K + 1, 2)
'           repl kodik("ep&aa with kodik("ep&epomeno
'           repl kodik("f&aa  with kodik("f&epomeno
'           repl kodik("p&aa  with kodik("p&epomeno
'       Next
'
'           aa = StrZero(15, 2)
'           repl kodik("ep&aa with ' '
'           repl kodik("f&aa with ' '
'           repl kodik("p&aa with '  '
'
'
'        ** ζωγραφίζω τα get
'        set color to &nw
'        for k=1 to kodik("nr
'            aa = StrZero(K, 2)
'            @ 1+k, 0 say k pict '99'
'            @ 1+k, 3 get kodik("ep&aa pict '@S20' valid find_field(kodik("ep&aa)=0
'            @ 1+k,25 get kodik("f&aa  pict '@S20'
'            @ 1+k,47 get kodik("p&aa  pict '@S20'
'            @ 1+k,70 get synola(k)    pict '9'
'        Next
'        SETCOLOR (LAST_COLOR)
'
'Return
'
'
'
'
'   Function MyErrorHandler(e, old)
'  *************************************************************
'      //
'   lOCAL cMessage:='',al(12)
'        // display message and traceback
'        if ( !Empty(e:osCode) )
'                cMessage += " (DOS Error " + LTRIM(e:osCode) + ") "
'        End
'
'         // build error message
'        cMessage := ErrorMessage(e)
'
'
'
'        // build options array
'        // aOptions := {"Break(Διακοπή)", "Quit"}
'        aOptions := {"Quit(Διακοπή)"}
'
'        if (e:canRetry)
'                AAdd(aOptions, "Retry(Ξαναπροσπαθώ)")
'        End
'
'        if (e:canDefault)
'                AAdd(aOptions, "Default(αγνοω-συνεχίζω)")
'        End
'
'
'        // put up alert box
'        nChoice := 0
'        while ( nChoice == 0 )
'
'                if ( Empty(e:osCode) )
'                        nChoice := Alert( cMessage, aOptions )
'
'                Else
'                        nChoice := Alert( cMessage + ;
'                                                        ";(DOS Error " + LTRIM(e:osCode) + ")", ;
'                                                        aOptions )
'                End
'
'
'                if ( nChoice == NIL )
'                End
'
'        End
'
'
'        if ( !Empty(nChoice) )
'
'                // do as instructed
'                if ( aOptions(nChoice) == "Break(Διακοπή)" )
'
'                elseif ( aOptions(nChoice) == "Quit(Διακοπή)" )
'                elseif ( aOptions(nChoice) == "Retry(Ξαναπροσπαθώ)" )
'                        return (.t.)
'
'                elseif ( aOptions(nChoice) == "Default(αγνοω-συνεχίζω)" )
'                        return (.f.)
'
'                End
'
'        End
'
'
'
'
'
'
'
'
'
'
'
'
'        Print cMessage
'        i := 2
'
'
'       * while ( !Empty(ProcName(i)) )
'       Print e: Description
'       For i = 1 To 12
'             if ( !Empty(ProcName(i)) )
'                al (i) = Trim(ProcName(i)) + "(" + Str(ProcLine(i)) + ")  "
'                Print al; (i)
'             Else
'                al(i)='....'
'             End If
'       Next
'
'
'     if .not. file('err.txt')
'       IF (nHandle := FCREATE('err.txt' , FC_NORMAL)) == -1
'         Print "File cannot be created:", FERROR()
'         QUIT
'       End If
'     Else
'       nHandle := FOPEN("err.txt", FO_READWRITE + FO_SHARED)
'       IF FERROR() != 0
'         Print "Cannot open file, DOS error ", FERROR()
'         QUIT
'       End If
'     End If
'
'
'       nLength := FSEEK(nHandle, 0, FS_END)
'
'      FWRITE(nHandle,dtoc(date())+' '+time()+' '+CMESSAGE +chr(10) )
'      For i = 1 To 12
'          if al(i)<>'....'
'             FWRITE(nHandle,'   '+al(i)+chr(10) )
'          End If
'      Next
'      FCLOSE (nHandle)
'       wait ''
'       BREAK      && objError          // Return error object to RECOVER
'     RETURN NIL
'
'
'
'
'
'
'
'
'
'
'
'Procedure opos_allo
'********************************
'Local GetList:={},sc
'priv m:=1,r
'R = readvar()
'sc = savescreen(24, 0, 24, 79)
'@ 24,0 say 'Οπως το πεδίο ' get m
'Read
'M = StrZero(M, 2)
'repl &r with kodik("f&m
'restscreen(24,0,24,79,sc)
'Return
'
'Function make_relations()
'**********************************************
'    if  len(mRelation2)>0
'       SELE &mARX2
'       if upper(indexkey(0)) = upper ( mrelation2 )
'          *
'       Else
'          * index on &mrelation2 to dok2&xeirisths
'       End If
'       sele &mARX1
'       set relation to &mrelation2 into &mARX2 ADDIT
'    End If
'
'    if  len(mRelation3)>0
'        set relation to &mrelation3 into &mARX3 ADDIT
'    End If
'
'    if len(mRelation4)>0
'        set relation to &mrelation4 into &mARX4 ADDIT
'    End If
'return nil
'
'
'proc check_kodik(mFile)
'****************************************************************
'priv strupin:={}
'AADD( STRUPIN , {"NUMBER_ARX","N",  1,0} )
'AADD( STRUPIN , {"ARX1      ","C", 40,0} )
'AADD( STRUPIN , {"ARX2      ","C", 40,0} )
'AADD( STRUPIN , {"ARX3      ","C", 40,0} )
'AADD( STRUPIN , {"ARX4      ","C", 40,0} )
'AADD( STRUPIN , {"RELATION4 ","C", 40,0} )
'AADD( STRUPIN , {"RELATION2 ","C", 40,0} )
'AADD( STRUPIN , {"RELATION3 ","C", 40,0} )
'AADD( STRUPIN , {"ART       ","C", 10,0} )
'AADD( STRUPIN , {"SYNOLO    ","C", 30,0} )
'AADD( STRUPIN , {"NO        ","C", 15,0} )
'AADD( STRUPIN , {"TAG       ","C", 60,0} )
'AADD( STRUPIN , {"EP01      ","C", 30,0} )
'AADD( STRUPIN , {"EP02      ","C", 30,0} )
'AADD( STRUPIN , {"EP03      ","C", 30,0} )
'AADD( STRUPIN , {"EP04      ","C", 30,0} )
'AADD( STRUPIN , {"EP05      ","C", 30,0} )
'AADD( STRUPIN , {"EP06      ","C", 35,0} )
'AADD( STRUPIN , {"EP07      ","C", 35,0} )
'AADD( STRUPIN , {"EP08      ","C", 30,0} )
'AADD( STRUPIN , {"EP09      ","C", 30,0} )
'AADD( STRUPIN , {"EP10      ","C", 30,0} )
'AADD( STRUPIN , {"EP11      ","C", 30,0} )
'AADD( STRUPIN , {"EP12      ","C", 30,0} )
'AADD( STRUPIN , {"EP13      ","C", 30,0} )
'AADD( STRUPIN , {"EP14      ","C", 30,0} )
'AADD( STRUPIN , {"EP15      ","C", 30,0} )
'AADD( STRUPIN , {"F01       ","C",180,0} )
'AADD( STRUPIN , {"F02       ","C",180,0} )
'AADD( STRUPIN , {"F03       ","C",180,0} )
'AADD( STRUPIN , {"F04       ","C",180,0} )
'AADD( STRUPIN , {"F05       ","C",180,0} )
'AADD( STRUPIN , {"F06       ","C",180,0} )
'AADD( STRUPIN , {"F07       ","C",180,0} )
'AADD( STRUPIN , {"F08       ","C",180,0} )
'AADD( STRUPIN , {"F09       ","C",180,0} )
'AADD( STRUPIN , {"F10       ","C",180,0} )
'AADD( STRUPIN , {"F11       ","C",180,0} )
'AADD( STRUPIN , {"F12       ","C",180,0} )
'AADD( STRUPIN , {"F13       ","C",180,0} )
'AADD( STRUPIN , {"F14       ","C",180,0} )
'AADD( STRUPIN , {"F15       ","C",180,0} )
'AADD( STRUPIN , {"P01       ","C", 30,0} )
'AADD( STRUPIN , {"P02       ","C", 30,0} )
'AADD( STRUPIN , {"P03       ","C", 30,0} )
'AADD( STRUPIN , {"P04       ","C", 30,0} )
'AADD( STRUPIN , {"P05       ","C", 30,0} )
'AADD( STRUPIN , {"P06       ","C", 30,0} )
'AADD( STRUPIN , {"P07       ","C", 30,0} )
'AADD( STRUPIN , {"P08       ","C", 30,0} )
'AADD( STRUPIN , {"P09       ","C", 30,0} )
'AADD( STRUPIN , {"P10       ","C", 30,0} )
'AADD( STRUPIN , {"P11       ","C", 30,0} )
'AADD( STRUPIN , {"P12       ","C", 60,0} )
'AADD( STRUPIN , {"P13       ","C", 30,0} )
'AADD( STRUPIN , {"P14       ","C", 30,0} )
'AADD( STRUPIN , {"P15       ","C", 30,0} )
'AADD( STRUPIN , {"SYNTWH    ","C", 80,0} )
'AADD( STRUPIN , {"SYNTIF    ","C",140,0} )
'AADD( STRUPIN , {"ARXEIO    ","C", 25,0} )
'AADD( STRUPIN , {"NR        ","N",  4,0} )
'AADD( STRUPIN , {"EPIKEF    ","C",150,0} )
'AADD( STRUPIN , {"S_EJECT   ","N",  2,0} )
'AADD( STRUPIN , {"FONTS     ","C", 30,0} )
'AADD( STRUPIN , {"K_HM1     ","C", 30,0} )
'AADD( STRUPIN , {"K_HM2     ","C", 30,0} )
'AADD( STRUPIN , {"K_CH1     ","C", 30,0} )
'AADD( STRUPIN , {"K_CH2     ","C", 30,0} )
'AADD( STRUPIN , {"K_CH3     ","C", 30,0} )
'AADD( STRUPIN , {"EP16      ","C", 50,0} )
'AADD( STRUPIN , {"EP17      ","C", 50,0} )
'AADD( STRUPIN , {"EP18      ","C", 50,0} )
'AADD( STRUPIN , {"EP19      ","C", 30,0} )
'AADD( STRUPIN , {"EP20      ","C", 30,0} )
'AADD( STRUPIN , {"F16       ","C",180,0} )
'AADD( STRUPIN , {"F17       ","C",180,0} )
'AADD( STRUPIN , {"F18       ","C",180,0} )
'AADD( STRUPIN , {"F19       ","C",180,0} )
'AADD( STRUPIN , {"F20       ","C",180,0} )
'AADD( STRUPIN , {"P16       ","C", 30,0} )
'AADD( STRUPIN , {"P17       ","C", 30,0} )
'AADD( STRUPIN , {"P18       ","C", 30,0} )
'AADD( STRUPIN , {"P19       ","C", 30,0} )
'AADD( STRUPIN , {"P20       ","C", 30,0} )
'AADD( STRUPIN , {"SYNT_ALAG ","C", 30,0} )
'AADD( STRUPIN , {"BSEIRA    ","C",160,0} )
'AADD( STRUPIN , {"SUM_SELTXT","C", 30,0} )
'AADD( STRUPIN , {"EK_METAF  ","N",  1,0} )
'AADD( STRUPIN , {"SE_METAF  ","N",  1,0} )
'AADD( STRUPIN , {"DIAX      ","C",  1,0} )
'AADD( STRUPIN , {"TOTAL     ","C",120,0} )
'AADD( STRUPIN , {"EP21      ","C", 30,0} )
'AADD( STRUPIN , {"EP22      ","C", 30,0} )
'AADD( STRUPIN , {"EP23      ","C", 30,0} )
'AADD( STRUPIN , {"EP24      ","C", 30,0} )
'AADD( STRUPIN , {"EP25      ","C", 30,0} )
'AADD( STRUPIN , {"P21       ","C", 30,0} )
'AADD( STRUPIN , {"P22       ","C", 30,0} )
'AADD( STRUPIN , {"P23       ","C", 30,0} )
'AADD( STRUPIN , {"P24       ","C", 30,0} )
'AADD( STRUPIN , {"P25       ","C", 30,0} )
'AADD( STRUPIN , {"F21       ","C",180,0} )
'AADD( STRUPIN , {"F22       ","C",180,0} )
'AADD( STRUPIN , {"F23       ","C",180,0} )
'AADD( STRUPIN , {"F24       ","C",180,0} )
'AADD( STRUPIN , {"F25       ","C",180,0} )
'AADD( STRUPIN , {"F26       ","C",180,0} )
'AADD( STRUPIN , {"F27       ","C",180,0} )
'AADD( STRUPIN , {"F28       ","C",180,0} )
'AADD( STRUPIN , {"F29       ","C",180,0} )
'AADD( STRUPIN , {"F30       ","C",180,0} )
'AADD( STRUPIN , {"EP26      ","C", 30,0} )
'AADD( STRUPIN , {"EP27      ","C", 30,0} )
'AADD( STRUPIN , {"EP28      ","C", 30,0} )
'AADD( STRUPIN , {"EP29      ","C", 30,0} )
'AADD( STRUPIN , {"EP30      ","C", 30,0} )
'AADD( STRUPIN , {"P26       ","C", 30,0} )
'AADD( STRUPIN , {"P27       ","C", 30,0} )
'AADD( STRUPIN , {"P28       ","C", 30,0} )
'AADD( STRUPIN , {"P29       ","C", 30,0} )
'AADD( STRUPIN , {"P30       ","C", 30,0} )
'AADD( STRUPIN , {"EPIKEF2   ","C", 70,0} )
'AADD( STRUPIN , {"EPIK_2    ","C", 70,0} )
'AADD( STRUPIN , {"EPIK_3    ","C", 70,0} )
'AADD( STRUPIN , {"K_CH4     ","C", 30,0} )
'AADD( STRUPIN , {"K_CH5     ","C", 30,0} )
'AADD( STRUPIN , {"PEDIA_TOT ","C", 80,0} )
'AADD( STRUPIN , {"FOR_TOT   ","C", 80,0} )
'AADD( STRUPIN , {"M_HEAD1   ","C", 40,0} )
'AADD( STRUPIN , {"M_HEAD2   ","C", 40,0} )
'AADD( STRUPIN , {"M_HEAD3   ","C", 40,0} )
'AADD( STRUPIN , {"M_HEAD4   ","C", 40,0} )
'
'lk_check (mFile)
'Return
'
'Procedure make_totals
'***********************************
'       if len(b)>1
'
'                  **************** κανω τις σούμες  **********************
'           if kodik("total<>'     '
'
'                 mtotal   =trim(kodik("total   )
'                 mfor_tot =trim(kodik("for_tot )
'
'                 for _kt=1 to 6
'                     _mTot(_kt) = subs( kodik("pedia_tot , 01+(_kt-1)*10 , 10 )
'                 Next
'
'
'                 _ST=0
'                 for _kt =1 to 6
'                     if _mTot(_kt)<>' '
'                         aa = alltrim ( str ( _kt ) )
'                        _st = _st + 1
'                         m&aa = _mTot(_kt)
'                     Else
'                        exit
'                     End If
'                 Next
'
'
'                 index on &mtotal to dok&xeirisths
'                 do case
'                    case _st=1
'                         total on  &b fields &m1 to dok&xeirisths for &mfor_tot
'                    case _st=2
'                         total on  &b fields &m1,&m2 to dok&xeirisths for &mfor_tot
'                    case _ST=3
'                         total on  &b fields &m1,&m2,&m3 to dok&xeirisths for &mfor_tot
'                    case _ST=4
'                         total on  &b fields &m1,&m2,&m3,&m4 to dok&xeirisths for &mfor_tot
'                    case _ST=5
'                         total on  &b fields &m1,&m2,&m3,&m4,&m5     to dok&xeirisths for &mfor_tot
'                    case _ST=6
'                         total on  &b fields &m1,&m2,&m3,&m4,&m5,&m6 to dok&xeirisths for &mfor_tot
'                    otherwise
'                         wait ' δεν όρισες πεδία που θα αθροίσω '
'                         Close Data
'                         Return
'                 endcase
'
'
'              use dok&xeirisths alias &marx1 exclu
'            * 20-2-2000    index on &b tag dok&xeirisths to dok&xeirisths
'
'              if  len(mRelation2)>0
'                  sele &mARX1
'                  set relation to &mrelation2 into &mARX2 ADDIT
'              End If
'              if  len(mRelation3)>0
'                  sele &mARX1
'                  set relation to &mrelation3 into &mARX3 ADDIT
'              End If
'              index on &b tag dok&xeirisths to dok&xeirisths
'
'
'           Else
'              Print b
'              Print xeirisths
'              Print dok
'              index on &b to dok___&xeirisths
'           End If
'           **************** κανω τις σούμες  **********************
'       End If
'
'Return
'
'
'proc make_mis_totals
'*********************************************************
'       if len(b)>1
'           gm_fILTER = TRIM(KODIK("SYNTIF)
'           if val(b)>0
'               set order to val(b)
'           Else
'               index on &b to dok&xeirisths for &gm_fILTER      && tag dok&xeirisths to dok&xeirisths
'           End If
'
'           **************** κανω τις σούμες  **********************
'           if kodik("total<>'     '
'              mtotal=trim(kodik("FOR_tot)
'              *  total on &b fields &mtotal  to dok&xeirisths
'
'
'
'              total on  &b fields     w1,w2,w3,w4,w5,w6,w7,w8,w9,w10,;
'                                      w11,w12,w13,w14,w15,w16,w17,w18,w19,w20,;
'                                      w21,w22,w23,w24,w25,w26,w27,w28,w29,w30,;
'                                      w31,w32,w33,w34,w35,w36,w37,w38,w39,w40,;
'                                      w41,w42,w43,w44,w45,w46,w47,w48,w49,w50,;
'                                      w51,w52,w53,w54,w55,w56,w57,w58,w59,w60,;
'                                      w61,w62,w63,w64,w65,w66,w67,w68,w69,w70,;
'                                      w71,w72,w73,w74,w75,w76,w77,w78,w79,w80,;
'                                      w81,w82,w83,w84,w85,w86,w87,w88,w89,w90,;
'                                      w91,w92,w93,w94,w95,w96,w97,w98,w99     ;
'                                  to dok&xeirisths for &mtotal
'              use dok&xeirisths alias &marx1 exclu
'              index on &b to dok&xeirisths   && tag dok&xeirisths to dok&xeirisths
'
'              if  len(mRelation2)>0
'                  SELE &mARX2
'                  set order to 1  &&  tag &mrelation2
'                  sele &mARX1
'                  set relation to &mrelation2 into &mARX2 ADDIT
'              End If
'           End If
'           **************** κανω τις σούμες  **********************
'
'       End If
'Return
'
'
'func errorMessage(e)
'local cMessage
'***********************************************
'
'        // start error message
'        cMessage := if( e:severity > ES_WARNING, "Error ", "Warning " )
'
'
'        // add subsystem name if available
'        if ( ValType(e:subsystem) == "C" )
'                cMessage += e:subsystem()
'        Else
'                cMessage += "???"
'        End
'
'
'        // add subsystem's error code if available
'        if ( ValType(e:subCode) == "N" )
'                cMessage += ("/" + LTRIM(STR(e:subCode) ))
'        Else
'                cMessage += "/???"
'        End
'
'
'        // add error description if available
'        if ( ValType(e:description) == "C" )
'                cMessage += ("  " + e:description)
'        End
'
'
'        // add either filename or operation
'        if ( !Empty(e:filename) )
'                cMessage += (": " + e:filename)
'
'        elseif ( !Empty(e:operation) )
'                cMessage += (": " + e:operation)
'
'        End
'
'
'return (cMessage)
'
'
'Function Find_timol(key)
'*****************************************
'* key   =  atim+dtos(hme)
'priv aa
'aa=select()
'do open_data with .f.,2,'EID','EGGTIM'
'
'sele EGGTIM
'
'seek key
'do while .not. eof() .and. key=eggtim("atim+dtos(eggtim("hme)
'  sele EID
'  seek eggtim("kode
'  sele EGGTIM
'  printer.print+1,15 say '- '+ eid("ono+'  '+trans(poso,'9999.9')+'  '+trans(timm,'9999,999.99')+'  '+trans(timm*poso,'999,999.99')
'  Skip
'enddo
'aa = alltrim(Str(aa))
'sele &aa
'
'Return '  '
'
'Function db17_udf()
'*****************************************
'para dbmode, fld
'getit=arr2(fld)
'
'do case
'  Case dbmode = 0
'     return 1
'  Case lastkey() = 27
'     return 0
'  Case lastkey() = K_INS
'       ADD_REC (0)
'       RETURN 1
'  Case lastkey() = K_ENTER
'     set cursor on
'     @ row(),col() get &getit
'     if rec_lock(0)
'       Read
'       unlock
'     End If
'     set cursor off
'     return 1
'  otherwise
'     return 1
'endcase
'
'
'
'
'
'
Function to928(string_ As String) As String
Dim A$, k As Integer, S As String, t As Integer, s928 As String, s437 As String
'metatrepei eggrafo apo 437->928
s928 = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ-αβγδεζηθικλμνξοπρστυφχψω-ςάέήίόύώ"
s437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰ-ªαβγεζηι" ' saehioyv
's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰª" 'αβγεζηι"


A$ = string_
'                                                        saehioyv
'GoTo 11
'Open Text2.Text For Output As #2
'Open Text1.Text For Input As #1
'Do While Not EOF(1)
'  Line Input #1, a$
  For k = 1 To Len(A$)
     S = Mid(A$, k, 1)
     t = InStr(s437, S)
     If t > 0 Then
        Mid$(A$, k, 1) = Mid$(s928, t, 1)
     End If
  Next
11
  to928 = A$


End Function

Function to437(string_ As String) As String
Dim A$, k As Integer, S As String, t As Integer, s928 As String, s437 As String
'metatrepei eggrafo apo 437->928
s928 = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ-αβγδεζηθικλμνξοπρστυφχψω-ςάέήίόύώ"
s437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰ-ªαβγεζηι" ' saehioyv
's437 = "€‚ƒ„…†‡‰‹‘’“”•–—-™› ΅Ά£¤¥¦§¨©«¬­®―ΰª" 'αβγεζηι"
'€‚ƒ„…†‡‰‹‘’“”•–—™› ΅Ά£¤¥¦§¨©«¬­®―ΰ

A$ = string_
'                                                        saehioyv
'GoTo 11
'Open Text2.Text For Output As #2
'Open Text1.Text For Input As #1
'Do While Not EOF(1)
'  Line Input #1, a$
  For k = 1 To Len(A$)
     S = Mid(A$, k, 1)
     t = InStr(s928, S)
     If t > 0 Then
        Mid$(A$, k, 1) = Mid$(s437, t, 1)
     End If
  Next
11
  to437 = A$


End Function

Sub g_ektyp(sql As String)
'POSITION EINAI TO RECORD ΤΟΥ REPORT
'notepad  hmm ok i checked the binary a bit out now..
'here are all command line arguments notepad takes
'
'/A <filename> open file as ansi
'/W <filename> open file as unicode
'/P <filename> print filename
'/PT <filename> <printername> <driverdll> <port>
'/.SETUP some weird stuff is happening i cant identify =) its enumerating the systemdir and opening notepad without the minimize option.. but dont ask me why
'
'hmm i think the best solution to your windowing problem would be to just move the window to where you want to have it
'close notepad so notepad saves the position size etc to HKEY_CURRENT_USER\Software\Microsoft\Notepad in the registry..
'then just export that registry.. create a batch file or something then that just imports that reg entry (regedit /s filename.reg) and opens your notepad after..

Dim DUM, n
Dim TT, F_T As New ADODB.Recordset


TT = GetCurrentTime()


F_T.Open sql, gdb, adOpenForwardOnly, adLockReadOnly

MDIForm1.Caption = (GetCurrentTime() - TT)


n = F_T.Fields.Count



f_tab(0) = 0
'f_tab(1) = 15
'f_tab(2) = 50
'f_tab(3) = 60
'f_tab(4) = 70

Dim k
Dim m_s

' m_s = f_tab(0)


'For K = 1 To f_t.Fields.Count - 1
'  f_tab(K) = 2 + m_s + f_t.Fields(K - 1).Size
'Next

'm_s = f_tab(0)








' Dim F_T As New ADODB.Recordset


TT = GetCurrentTime()


F_T.Open sql, gdb, adOpenForwardOnly, adLockReadOnly

n = F_T.Fields.Count - 1



For k = 1 To n
    If F_T(k - 1).Type = 8 Then 'DATE
      f_tab(k) = 2 + f_tab(k - 1) + 10
    Else
      f_tab(k) = 2 + f_tab(k - 1) + F_T.Fields(k - 1).DefinedSize
    End If
Next


'dum = print3_xar(n, "0", "κωδικός", "F_KODIK!FIELDS", "xxxxxxx")








End Sub




