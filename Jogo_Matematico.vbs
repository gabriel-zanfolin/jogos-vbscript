dim operando1,operando2,operador,resultado,resp1,resp2,qtde_acertos,recorde,audio
recorde = 0
qtde_acertos = 0
resultado = 1
resp1 = 1
call carregar_audio

sub carregar_audio()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate = 3 'Velocidade da fala
call inicio
end sub

sub inicio()
do while (resultado=resp1)
	call sorteia_operando1
loop
end sub

sub sorteia_operando1()
randomize(second(time))
operando1=int(rnd * 10) + 1
call sorteia_operando2
end sub

sub sorteia_operando2()
randomize(second(time))
operando2=int(rnd * 10) + 1
call sorteia_operador
end sub

sub sorteia_operador()
randomize(second(time))
operador=int(rnd * 3) + 1
select case operador
case 1:
	operador = "+"
	resultado = operando1 + operando2
case 2:
	operador = "-"
	resultado = operando1 - operando2
case 3:
	operador = "*"
	resultado = operando1 * operando2
end select
call equacao
end sub

sub equacao()
resp1=cint(inputbox("=================================" + vbnewline &_
"   ACERTE O CÁLCULO MATEMÁTICO   " + vbnewline &_
"=================================" + vbnewline &_
"	RESOLVA:   "& operando1 &" "& operador &" "& operando2 &" = ???" + vbnewline &_
"=================================", "SISTEMAS DE INFORMAÇÃO - JOGO MATEMÁTICA"))
if (resp1=resultado) then
	qtde_acertos = qtde_acertos + 1
	if (qtde_acertos>recorde) then
		recorde = qtde_acertos
	end if
	call msg_acerto
else
	call fim_jogo
end if
end sub

sub msg_acerto()
audio.speak ("Parabéns, você ACERTOU!!!")
msgbox("  Parabéns, você ACERTOU!!!" + vbnewline &_
       "  Quantidade de Acertos: "& qtde_acertos &"" + vbnewline &_
       "===================" + vbnewline &_
       "	Recorde: "& recorde &"" + vbnewline &_
       "==================="), vbokonly + vbinformation, "AVISO"
call sorteia_operando1
end sub

sub fim_jogo()
audio.speak ("Que pena, você ERROU!!!")
msgbox("  Que pena, você ERROU!!!" + vbnewline &_
       "  Quantidade de Acertos: "& qtde_acertos &"" + vbnewline &_
       "===================" + vbnewline &_
       "	Recorde: "& recorde &"" + vbnewline &_
       "==================="), vbokonly + vbcritical, "AVISO"
resp2=msgbox("Deseja jogar novamente?", vbyesno + vbquestion, "AVISO")
if (resp2=vbyes) then
	qtde_acertos = 0
	call sorteia_operando1
else
	wscript.quit
end if
end sub