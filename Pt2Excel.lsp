;-------------------------------------------------------------------------------
; Program Name: Pt2Excel.lsp [Pt2Excel R0]
; Created By:   Max Rocha (email: max.warocha@gmail.com)
; Date Created: 4-12-2021
; Function:     Opens  a  new  Excel  spreadsheet  and  write  the  coordinates
;			    of selected points 
; Note:         It's required that users load Terry Miller's GetExcel.lsp before
;               calling Pt2Excel.(https://autolisp-exchange.com/LISP/GetExcel.lsp)
;-------------------------------------------------------------------------------
; Revision History
; Rev  By     Date    Description
;-------------------------------------------------------------------------------
; 0    MR   4-12-21   Initial version
;-------------------------------------------------------------------------------

(defun C:Pt2Excel ()

	;CHECK IF GETEXCEL IS LOADED
	(if (= (if getexcel T nil) nil) (progn (alert "GETEXCEL NOT LOADED!\n GETEXCEL NAO CARREGADO!") (vl-exit-with-error "")))

	;OPEN A NEW EXCEL INSTANCE
	(openexcel nil (strcat (getvar "dwgname") " PTs") t)

	;LANGUAGE (PT-BR or EN-US - DEFAULT: PT-BR)
	;Selecting a language
	(initget "Portugues English")
	(setq language (cond ((getkword "\nPT2EXCEL - Please choose your language: [Portugues/English] <Portugues>")) ("Portuguese")))
		
	;PORTUGUESE TEXTs
	(If (= language "Portugues")
	
		(progn
		
			(setq textYes "Sim")
			(setq textNo "Nao")
			(setq textPoint "Ponto")
			(setq textComma "Virgula")
			
			(setq textDecimalSeparator (strcat "\nPT2EXCEL - Selecione um separador decimal: [" textComma "/" textPoint "] <Virgula>"))
			(setq textPointFormat "\nPT2EXCEL - Selecione o formato de saida dos pontos: [PENZ/PNEZ/PENZD/PNEZD] <PENZ>")
			(setq textCreatePoint (strcat "\nPT2EXCEL - Criar ponto ao clicar? [" textYes "/" textNo "] <Nao>"))
			(setq textPointPrefix "\nPT2EXCEL - Insira o prefixo de designacao dos pontos: (ex. P => P1)")
			(setq textSelectAPoint "\nPT2EXCEL - Clique em um ponto: ")
			(setq textGetDescription "\nPT2EXCEL - Insira a descricao do ponto: ")	
		)
	)
	
	;ENGLISH TEXTs
	(If (= language "English")
	
		(progn
		
			(setq textYes "Yes")
			(setq textNo "No")
			(setq textPoint "Point")
			(setq textComma "Comma")
			
			(setq textDecimalSeparator (strcat "\nPT2EXCEL - Select the desired decimal separator: [" textComma "/" textPoint "] <Comma>"))
			(setq textPointFormat "\nPT2EXCEL - Choose point output format: [PENZ/PNEZ/PENZD/PNEZD] <PENZ>")
			(setq textCreatePoint (strcat "\nPT2EXCEL - Draw a new point after clicking? [" textYes "/" textNo "] <No>"))
			(setq textPointPrefix "\nPT2EXCEL - Insert the desired point prefix: (e.g. P => P1)")
			(setq textSelectAPoint "\nPT2EXCEL - Click in the desired point: ")
			(setq textGetDescription "\nPT2EXCEL - Insert point description: ")	
		)
	)
		
	;SETUP
	;Select a decimal separator
	(initget (strcat textComma " " textPoint))
	(setq decimalSeparator (cond ((getkword textDecimalSeparator)) (textComma)))
	(if (= decimalSeparator textComma) (setq decimalSeparator ","))
	(if (= decimalSeparator textPoint) (setq decimalSeparator "."))

	;Select output point format
	(initget "PENZ PNEZ PENZD PNEZD")
	(setq pointFormat (cond ((getkword textPointFormat)) ("PENZ")))

	;Create a point entity after clicking?
	(initget (strcat textYes " " textNo))
	(setq createPoint (cond ((getkword textCreatePoint)) (textNo)))
	
	;Points prefix
	(setq prefix (getstring textPointPrefix))
	
	;COUNTER
	(setq i 1)
	
	;FILLING THE FIRST LINE OF THE SPREADSHEET
	;1st column
	(PutCell (strcat "A" (itoa i)) "P")
	
	;2nd and 3rd columns (EN or NE)
	(if (or (= pointformat "PENZ") (= pointFormat "PENZD"))

		(progn
	
			(PutCell (strcat "B" (itoa i)) "E")
			(PutCell (strcat "C" (itoa i)) "N")
		)
	)

	(if (or (= pointformat "PNEZ") (= pointFormat "PNEZD"))

		(progn
	
			(PutCell (strcat "B" (itoa i)) "N")
			(PutCell (strcat "C" (itoa i)) "E")
		)
	)

	;4th column
	(PutCell (strcat "D" (itoa i)) "Z")

	;5th column
	(if (or (= pointFormat "PENZD") (= pointFormat "PNEZD")) (PutCell (strcat "E" (itoa i)) "D"))

	;MAIN CODE
	;GetPoints
	(while (setq pt (getpoint textSelectAPoint))
		
		;Draw the point
		(if (= createPoint textYes) (command "_.point" pt))
			
		;Get point description
		(if (or (= pointFormat "PENZD") (= pointFormat "PNEZD")) (setq description (getstring textGetDescription)))
			
		;Filling excel i+1th line
		;1st column
		(PutCell (strcat "A" (itoa (+ i 1))) (strcat prefix (itoa i)))

		;2nd and 3rd columns (EN or NE) - Also substitutes . => selected decimal separator 
		(if (or (= pointFormat "PENZ") (= pointFormat "PENZD"))
					
			(progn 		
					
				(PutCell (strcat "B" (itoa (+ i 1))) (vl-string-subst decimalseparator "." (rtos (car pt))))
				(PutCell (strcat "C" (itoa (+ i 1))) (vl-string-subst decimalseparator "." (rtos (cadr pt))))
						
			)
		)
				
		(if (or (= pointFormat "PNEZ") (= pointFormat "PNEZD"))
					
			(progn 		
					
				(PutCell (strcat "B" (itoa (+ i 1))) (vl-string-subst decimalseparator "." (rtos (cadr pt))))
				(PutCell (strcat "C" (itoa (+ i 1))) (vl-string-subst decimalseparator "." (rtos (car pt))))
						
			)
		)
			
		;4th column
		(PutCell (strcat "D" (itoa (+ i 1))) (vl-string-subst decimalseparator "." (rtos (caddr pt))))
			
		;5th column
		(PutCell (strcat "E" (itoa (+ i 1))) description)
			
		;Increases counter by 1
		(setq i (+ i 1))
	)
)