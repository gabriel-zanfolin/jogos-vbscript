dim nome,nvl1(5),nvl2(5),nvl3(5),nvl4(5),qtde_acertos,replay,cont,nvl,aux
dim pulo,audio,palavra,n_vetor,n_ultimo,n_penultimo,n_antipenultimo
dim palpite,pulada,x,y
qtde_acertos=0
pulada=0
aux=0
x=1
y=1
pulo=1
n_penultimo=0
n_ultimo=0
n_vetor=0
call carregar_audio

sub carregar_audio()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate = 2 'Velocidade da fala
call carregar_palavras
end sub

sub carregar_palavras()
'N�vel 1
nvl1(1) = "GAFANHOTO"
nvl1(2) = "UMBIGO"
nvl1(3) = "VIDRO"
nvl1(4) = "FAROFA"
nvl1(5) = "X�CARA"
'N�vel 2
nvl2(1) = "TR�PEDE"
nvl2(2) = "ESPALHAFATOSO"
nvl2(3) = "QU�NTICO"
nvl2(4) = "SOBRANCELHA"
nvl2(5) = "CABELEIREIRO"
'N�vel 3
nvl3(1) = "PARALELEP�PEDO"
nvl3(2) = "COINCID�NCIA"
nvl3(3) = "RETR�GRADO"
nvl3(4) = "METEREOLOGIA"
nvl3(5) = "DELINQUENTE"
'N�vel 4
nvl4(1) = "QUINQUILHARIA"
nvl4(2) = "INTERDISCIPLINIDADE"
nvl4(3) = "IMPEACHMENT"
nvl4(4) = "VICISSITUDE"
nvl4(5) = "OTORRINOLARINGOLOGISTA"
call insere_nome
end sub

sub insere_nome()
nome=cstr(inputbox("Digite o nome do jogador...", "SOLETRANDO"))
call sorteia_palavra
end sub

sub sorteia_palavra()
for nvl=x to 4 step 1
	if (aux=0) then
		y = 1
	else
		y = cont
	end if
	for cont=y to 3 step 1
		n_antipenultimo = n_penultimo
		n_penultimo = n_ultimo
		n_ultimo = n_vetor
		if (cont=1) then
			randomize(second(time))
			n_vetor=int(rnd * 5) + 1
			do while (n_vetor=n_ultimo)
				randomize(second(time))
				n_vetor=int(rnd * 5) + 1
			loop
			replay = 1
			call dita_palavra
		else
			if (cont=2) then
				randomize(second(time))
				n_vetor=int(rnd * 5) + 1
				do while (n_vetor=n_ultimo or n_vetor=n_penultimo)
					randomize(second(time))
					n_vetor=int(rnd * 5) + 1
				loop
				replay = 1
				call dita_palavra
			else
				if (cont=3) then
					randomize(second(time))
					n_vetor=int(rnd * 5) + 1
					do while (n_vetor=n_ultimo or n_vetor=n_penultimo or n_vetor=n_antipenultimo)
						randomize(second(time))
						n_vetor=int(rnd * 5) + 1
					loop
					replay = 1
					call dita_palavra
				end if
			end if
		end if
		aux = 0
	next
	x = x + 1
next
call msg_vitoria
end sub

sub dita_palavra()
if (nvl=1) then
	palavra=nvl1(n_vetor)
else
	if (nvl=2) then
		palavra=nvl2(n_vetor)
	else
		if (nvl=3) then
			palavra=nvl3(n_vetor)
		else
			palavra=nvl4(n_vetor)
		end if
	end if
end if
if (palavra=pulada) then
	call sorteia_palavra
end if
audio.speak (palavra)
call insere_tentativa
end sub

sub insere_tentativa()
if (replay=1 and pulo=1) then
	palpite=cstr(inputbox("DIGITE A PALAVRA OUVIDA" + vbnewline + vbnewline &_
			      "Nome do Jogador: "& nome &"" + vbnewline + vbnewline &_
			      "================================" + vbnewline &_
			      "[O]uvir Novamente a Palavra" + vbnewline &_
			      "[P]ular a Palavra uma �nica vez" + vbnewline &_
			      "================================", "SOLETRANDO"))
	palpite = UCase(palpite)
	if (palpite=palavra) then
		call msg_acerto
	else
		if (palpite="O") then
			replay = 0
			call dita_palavra
		else
			if (palpite="P") then
				aux = y
				pulo = 0
				pulada = palavra
				call sorteia_palavra
			else
				call msg_erro
			end if
		end if
	end if
else
	if (replay=0 and pulo=1) then
		palpite=cstr(inputbox("DIGITE A PALAVRA OUVIDA" + vbnewline + vbnewline &_
				      "Nome do Jogador: "& nome &"" + vbnewline + vbnewline &_
				      "================================" + vbnewline &_
				      "[P]ular a Palavra uma �nica vez" + vbnewline &_
				      "================================", "SOLETRANDO"))
		palpite = UCase(palpite)
		if (palpite=palavra) then
			call msg_acerto
		else
			if (palpite="P") then
				aux = y
				pulo = 0
				pulada = palavra
				call sorteia_palavra
			else
				call msg_erro
			end if
		end if
	else
		if (replay=1 and pulo=0) then
			palpite=cstr(inputbox("DIGITE A PALAVRA OUVIDA" + vbnewline + vbnewline &_
				      "Nome do Jogador: "& nome &"" + vbnewline + vbnewline &_
				      "================================" + vbnewline &_
			      	      "[O]uvir Novamente a Palavra" + vbnewline &_
			              "================================", "SOLETRANDO"))
			palpite = UCase(palpite)
			if (palpite=palavra) then
				call msg_acerto
			else
				if (palpite="O") then
					replay = 0
					call dita_palavra
				else
					call msg_erro
				end if
			end if
		else
			if (replay=0 and pulo=0) then
				palpite=cstr(inputbox("DIGITE A PALAVRA OUVIDA" + vbnewline + vbnewline &_
			      			      "Nome do Jogador: "& nome &"", "SOLETRANDO"))
				palpite = UCase(palpite)
				if (palpite=palavra) then
					call msg_acerto
				else
					call msg_erro
				end if
			end if
		end if
	end if
end if
end sub

sub msg_acerto()
qtde_acertos = qtde_acertos + 1
msgbox("Parab�ns "& nome &", voc� acertou!" + vbnewline &_
       "Qtde de Acertos: "& qtde_acertos &" de 12" + vbnewline &_
       "N�vel 0"& nvl &""), vbinformation + vbokonly, "AVISO"
end sub

sub msg_erro()
msgbox("Que pena "& nome &", voc� errou!" + vbnewline &_
       "Qtde de Acertos: "& qtde_acertos &"" + vbnewline &_
       "N�vel 0"& nvl &""), vbcritical + vbokonly, "ATEN��O"
resp=msgbox("Deseja jogar novamente?", vbyesno + vbquestion, "AVISO")
if (resp=vbyes) then
	qtde_acertos=0
	pulada=0
	aux=0
	x=1
	y=1
	pulo=1
	n_penultimo=0
	n_ultimo=0
	n_vetor=0
	call sorteia_palavra
else
	wscript.quit
end if
end sub

sub msg_vitoria()
msgbox("==============================" + vbnewline &_
       "      PARAB�NS, VOC� GANHOU O JOGO!!!" + vbnewline &_
       "=============================="), vbinformation + vbokonly, "AVISO"
resp=msgbox("Deseja jogar novamente?", vbyesno + vbquestion, "AVISO")
if (resp=vbyes) then
	qtde_acertos=0
	pulada=0
	aux=0
	x=1
	y=1
	pulo=1
	n_penultimo=0
	n_ultimo=0
	n_vetor=0
	call sorteia_palavra
else
	wscript.quit
end if
end sub
