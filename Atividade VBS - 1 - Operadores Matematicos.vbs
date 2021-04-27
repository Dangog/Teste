dim n1,n2,op,resp,result,opc,cont

call zera_cont

sub zera_cont	
cont=0
call inicio
end sub

call inicio
sub inicio()

randomize(second(time)) 
	op=int(rnd * 3) + 1

randomize(second(time)) 
	n1=int(rnd * 10) + 1

randomize(second(time)) 
	n2=int(rnd * 10) + 1

if op = 1 then
	result = n1 - n2 
	opc = "-"
	
	elseif op = 2 then
		result = n1 + n2 
		opc = "+"
		
			elseif op = 3 then
			opc = "*"
			result = n1 * n2 

end if

resp=cint(inputbox("=============================" + vbnewline &_
		"ACERTE O CALCULO MATEMÁTICO" + vbnewline &_
		"============================" + vbnewline &_
		""& n1 &" "& opc &" "& n2 &"= ???" + vbnewline &_
		"=============================","RESOLVA"))
		if resp = result then
			call ganhou_sub
				elseif resp <> result then
				call perdeu_sub	
		end if
end sub

sub ganhou_sub()
cont = cont+1
msgbox("PARABÉNS VOCÊ GANHOU" + vbnewline &_
		"Qtde de acertos: "& cont &""),vbexclamation+vbokonly,"AVISO"
		call inicio
end sub
	
sub perdeu_sub()
msgbox("Você errou!" + vbnewline &_
		"Qtde de acertos: "& cont &""),vbCritical,"ATENÇÃO"
		
resp=msgbox("Deseja jogar novamente?",vbquestion+vbYesNo,"ATENÇÃO")
	if resp = vbyes then
		call zera_cont
			else	
			wscript.quit
end if

resp=msgbox("Deseja jogar novamente?",vbquestion+vbYesNo,"ATENÇÃO")
	if resp = vbyes then
		call zera_cont
			else	
			wscript.quit
end if

resp=msgbox("Deseja jogar novamente?",vbquestion+vbYesNo,"ATENÇÃO")
	if resp = vbyes then
		call zera_cont
			else	
			wscript.quit
			
	resp=msgbox("Deseja jogar novamente?",vbquestion+vbYesNo,"ATENÇÃO")
	if resp = vbyes then
		call zera_cont
			else	
			wscript.quit
end if
end if
end sub


