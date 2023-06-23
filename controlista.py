#coding: utf-8
import visa
import time
import datetime
import pyperclip
import math
import numpy as np
import keyboard
import ctypes
import statistics
import nidaqmx
from thermocouples_reference import thermocouples

import colorama
from colorama import Fore, Back, Style
colorama.init()
# colorama.init(autoreset=True)

#from termcolor import colored

#########################################################################
# funções de trabalho
#########################################################################

def listaInstrumentos(endGpib='0'):
	rm = visa.ResourceManager()
	instrumentos = rm.list_resources()
	
	print('\n')
	print('Instrumentos conectados:\n')

	for instrumento in instrumentos:
				
		if instrumento != instrumentos[0]:
			print ('- '+instrumento)
			x = rm.open_resource(instrumento, timeout=1000, read_termination='\n')
			
			try:
				q = x.query('ID?') #pro 3458A
				x.query('ERR?')
			
			except Exception as e:
				
				try:
					q = x.query('*IDN?')
			
				except Exception as e:
					q = 'não respondeu o nome'

			x.control_ren(6)
			print(' > '+q+'\n')

def pegaEnderecos():
	
	endGpib = input('Qual interface GPIB? \n> ')

	if endGpib[0].lower() == 'n':
		listaInstrumentos()
		endGpib = input('Qual interface GPIB? \n> ')

	endereco = input('Qual endereço GPIB?\n> ')
	if endereco[0].lower() == 'n':
		listaInstrumentos()
		endereco = input('Qual endereço GPIB?\n> ')
	
	return endGpib, endereco

def defineSetup(ajustaCasas = False, unidadeSelecionada = False, nplc = False):

	print('\n> Quantas leituras?')
	quantasLeituras = str(int(input('> ')))
	if quantasLeituras == 'voltar':
		main()

	if nplc == True:
		print('> Qual NPLC? (normalmente 100)')
		nplc = input('> ')

	if ajustaCasas == True:
		print('\n> Qual multiplicador no Autolab?')
		print('p | n | µ | m | nenhum | k | M | G')
		print('> [última pergunta]')
		ajustaCasas = input('> ')

		if ajustaCasas == 'pico' or ajustaCasas == 'p':
			unidadeSelecionada = 'pico'
			ajustaCasas = 1/1_000_000_000_000

		if ajustaCasas == 'nano' or ajustaCasas == 'n':
			unidadeSelecionada = 'nano'
			ajustaCasas = 1/1_000_000_000

		if ajustaCasas == 'micro' or ajustaCasas == 'u':
			unidadeSelecionada = 'micro'
			ajustaCasas = 1/1_000_000

		if ajustaCasas == 'mili' or ajustaCasas == 'm':
			unidadeSelecionada = 'mili'
			ajustaCasas = 1/1_000

		if ajustaCasas == 'nenhum' or ajustaCasas == '':
			unidadeSelecionada = 'sem mult.'
			ajustaCasas = 1

		if ajustaCasas == 'kilo' or ajustaCasas == 'k':
			unidadeSelecionada = 'kilo'
			ajustaCasas = 1_000

		if ajustaCasas == 'mega' or ajustaCasas == 'M':
			unidadeSelecionada = 'mega'
			ajustaCasas = 1_000_000

		if ajustaCasas == 'giga' or ajustaCasas == 'G':
			unidadeSelecionada = 'giga'
			ajustaCasas = 1_000_000_000

	

	setup = {
		'quantasLeituras': quantasLeituras,
		'unidadeSelecionada': unidadeSelecionada,
		'ajustaCasas': ajustaCasas,
		'nplc': nplc
	}

	return setup

def processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto):

	a = str(round(a/float(ajustaCasas),12))
	a = a.replace('.', ',')
	indice = str(leitura+1)
	indice = '[{}/{}]'.format(indice, quantasLeituras)
	print('{}\t{}'.format(indice, a))
	leiturasPonto_txt.write(a+'\n')
	leituras_txt.write(a+'\n')
	setPonto.append(a)

def colaLeiturasNoExcel(endGpib):

	if endGpib == '0':
		keyboard.press_and_release('ctrl + 0')

	if endGpib != '0':
		keyboard.press_and_release('ctrl + 9')
	
def proximoPonto(leiturasPonto_txt, leituras_txt, x, endGpib, temReset=False):

	mesmoSetup = 'null'
	resetFuncao = False


	leiturasPonto_txt.close()
	leituras_txt.write('\n')

	if x != 'placa NI':
		x.control_ren(6)
	
	colaLeiturasNoExcel(endGpib)

	print('\n> Próximo ponto? s/n/exp')
	a = input('> ').lower()
	if a == '':
		a = 's'

	while a[0] == 'e':
		setPontos = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'r', encoding='utf-8')
		pyperclip.copy(str(setPontos.read()))
		setPontos.close()
		print('Leituras copiadas!')
		print('\n> Próximo ponto? s/n')
		a = input('> ').lower()
		if a == '':
			a = 's'

	
	if a == 'voltar' or a == 'n':
		leituras_txt.close()
		if temReset == True:
			resetFuncao = True
		else:
			main()

	else:
		print('\n> Mesma configuração? s/n')
		mesmoSetup = input('> ').lower()

	
	if temReset == True:
		saida = {
			'mesmoSetup': mesmoSetup,
			'resetFuncao': resetFuncao
		}
	else:
		saida = mesmoSetup
	
	return saida


#########################################################################
# 5700AsII
#########################################################################

def cal5700AsII(endGpib, endereco):
	
	# conecta
	rm = visa.ResourceManager()
	cal5700AsII = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco))
	
	#pré-setup
	cal5700AsII.write('out 0V, 0 Hz; stby')
	cal5700AsII.write('*SRE 32; *ESE 1; *cls')

	print('Qual função?')
	print('v | i | ph | tc | livre')
	funcao = input('> ').lower()

	funcoes = {
		'v': 'Quantos Volts? [x V | x V, x Hz]',
		'i': 'Quantos amperes? [x mA | x A, x Hz]',
		'ph': 'Qual pH?',
		'tc': 'Qual temperatura? (em °C)',
		'livre': 'Digite o comando completo a ser enviado',
	}

	if funcao == 'ph':
		print('Simular pH para qual temperatura?')
		temperatura = input('> ')

	if funcao == 'tc':
		print('Simular tensão para qual termopar?')
		qual_tc = input('> ').upper()
		tc = thermocouples[qual_tc]
		print('Qual temperatura ambiente? (em °C)')
		tempAmb = input('> ')


	while True:
		
		print('\n'+funcoes[funcao])
		a = input('> ')
		

		if a == 'ex grd on' or a == 'exgrd on' or a == 'extguard on':
			cal5700AsII.write('extguard on')

		if a == 'ex grd off' or a == 'exgrd off' or a == 'extguard off':
			cal5700AsII.write('extguard off')

		
		if a == "":
			a = "stby"

		if a[-1] != 'z':
			a = a.replace(',', '.')

		if a[0:3] == 'vol':
			cal5700AsII.write('out 0V, 0 Hz; stby')
			break

		if a[-1] == 'p':
			a_len = len(a)
			if a.find(' vpp') > -1:
				a = a[0:a_len-4] # removendo o espaço e o "vpp"
			else:
				a = a[0:a_len-3] # removendo só "vpp"
			a = str(round(float(a) /( 2 * math.sqrt(2)), 7) )
		
		if a[0:3] == 'stb':
			cal5700AsII.write('stby')
			
		elif a[-1] == 'z' :
				a = a.upper().replace('HZ', 'Hz').replace('K', 'k')
				cal5700AsII.write('*SRE 32; *ESE 1; *cls')
				cal5700AsII.write('out {}; oper'.format(a))
				print('out {}; oper'.format(a))
		else:
			
			if funcao == 'v':
				cal5700AsII.write('*SRE 32; *ESE 1; *cls')
				cal5700AsII.write('out {}V; oper'.format(a))
				print('out {}V; oper'.format(a))

			if funcao == 'i':
				a = a.lower()
				unidade = 'mA'
				unidadePrint = 'mA'
				
				if a.find('n') > -1:
					a = a.replace('na', '')
					a = a.replace('n', '')
					unidade = 'e-9A'
					unidadePrint = 'nA'
				
				if a.find('u') > -1:
					a = a.replace('uA', '')
					a = a.replace('u', '')
					unidade = 'e-6A'
					unidadePrint = 'uA'

				if a.find('m') > -1:
					a = a.replace('ma', '')
					a = a.replace('m', '')
					unidade = 'e-3A'
					unidadePrint = 'mA'

					
				if a.find('a') > -1 and a.find('u') == -1 and a.find('m') == -1:
					a = a.replace('a', '')
					unidade = 'A'
					unidadePrint = 'A'

				cal5700AsII.write('out {}{}; oper'.format(a, unidade))
				print('out {} {}; oper'.format(a, unidadePrint))

			if funcao == 'ph':

				if float(a) > 20 or float(a) < -2 :
					cal5700AsII.write('stby')
					print('Opa! Máx: 20 pH | Mín: -2 pH')

				else:
					vph = float(a)
					temperatura = float(temperatura)
					F = 96485.33289 #constante de Faraday
					R = 8.3144598 #constante universal dos gases

					# vph = 7-(a*F/(math.log(10)*R*(273.15+temperatura)))
					# (a*F/(math.log(10)*R*(273.15+temperatura))) + vph = 7
					# a*F/(math.log(10)*R*(273.15+temperatura)) = 7 - vph
					# a*F = (7 - vph) * (math.log(10)*R*(273.15+temperatura))
					a = ((7 - vph) * (math.log(10)*R*(273.15+temperatura)))/F
					a = round(a,9)
					
					cal5700AsII.write('out {}V; oper'.format(a))
					print('out {}V; oper'.format(a))

			if funcao == 'tc':
				# tensaoDesejada = tc.emf_mVC(float(a), Tref = 0)
				# tensaoTempAmbiente = tc.emf_mVC(float(tempAmb), Tref = 0)
				# tensaoProCalibrador = tensaoDesejada - tensaoTempAmbiente
				tensaoProCalibrador = tc.emf_mVC(float(a), Tref = float(tempAmb))
				assert abs(tensaoProCalibrador) < 150
				print('out {}E-3V; oper'.format(tensaoProCalibrador))
				cal5700AsII.write('out {}E-3V; oper'.format(tensaoProCalibrador))
			
			if funcao == 'livre':
				cal5700AsII.write('{}'.format(a))

#########################################################################
# 5520A
#########################################################################

def cal55XXX(endGpib, endereco):
	
	# conecta
	rm = visa.ResourceManager()
	cal55XXX = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco))
	
	#pré-setup
	cal55XXX.write('stby')
	cal55XXX.write('*SRE 32; *ESE 1; *cls')
	
	print('Qual função?')
	print('v | i | ohm | cap | tc | tempo | livre')
	funcao = input('> ').lower()

	funcoes = {
		'v': 'Quantos Volts? [x V | x V, x Hz]',
		'i': 'Quantos amperes? [x mA | x A, x Hz]',
		'ohm': 'Quantos Ohms?',
		'cap': 'Quantos Farad?',
		'tc': 'Qual temperatura? (em °C)',
		'tempo': 'Quantos segundos?',
		'livre': 'Digite o comando completo a ser enviado',
	}

	if funcao == 'tc':
		print('Simular tensão para qual termopar?')
		qual_tc = input('> ').upper()
		tc = thermocouples[qual_tc]
		print('Qual temperatura ambiente? (em °C)')
		tempAmb = input('> ')

	while True:
		
		print('\n'+funcoes[funcao])
		a = input('> ')
	

		if a == "":
			a = "stby"

		if a[-1] != 'z':
			a = a.replace(',', '.')

		if a[0:3] == 'vol':
			cal55XXX.write('out 0V, 0 Hz; stby')
			break
		
		if a[0:3] == 'stb':
			cal55XXX.write('stby')
			
		elif a[-1] == 'z':
			a = a.upper().replace('HZ', 'Hz').replace('K', 'k')
			
			cal55XXX.write('*SRE 32; *ESE 1; *cls')
			cal55XXX.write('out {}; oper'.format(a))
			print('out {}; oper'.format(a))
		else:
			
			if funcao == 'v':

				if a[-1] == 'p':
					a_len = len(a)
					if a.find(' vpp') > -1:
						a = a[0:a_len-4] # removendo o espaço e o "vpp"
					else:
						a = a[0:a_len-3] # removendo só "vpp"
					a = str(round(float(a) /( 2 * math.sqrt(2)), 7) )

				
				cal55XXX.write('*SRE 32; *ESE 1; *cls')
				cal55XXX.write('out {}V; oper'.format(a))
				print('out {}V; oper'.format(a))

			if funcao == 'i':
				a = a.lower()
				unidade = 'mA'
				unidadePrint = 'mA'
				
				if a.find('n') > -1:
					a = a.replace('na', '')
					a = a.replace('n', '')
					unidade = 'e-9A'
					unidadePrint = 'nA'
				
				if a.find('u') > -1:
					a = a.replace('uA', '')
					a = a.replace('u', '')
					unidade = 'e-6A'
					unidadePrint = 'uA'

				if a.find('m') > -1:
					a = a.replace('ma', '')
					a = a.replace('m', '')
					unidade = 'e-3A'
					unidadePrint = 'mA'

					
				if a.find('a') > -1 and a.find('u') == -1 and a.find('m') == -1:
					a = a.replace('a', '')
					unidade = 'A'
					unidadePrint = 'A'

				cal55XXX.write('out {}{}; oper'.format(a, unidade))
				print('out {} {}; oper'.format(a, unidadePrint))

			if funcao == 'ohm':
				cal55XXX.write('out {}Ohm; oper'.format(a))
				print('out {} Ohm; oper'.format(a))

			if funcao == 'cap':
				if a.lower() == '0 ohm':
					cal55XXX.write('out 0 Ohm; oper')
					print('out 0 Ohm; oper')
				else:
					cal55XXX.write('out {}F; oper'.format(a))
					print('out {}F; oper'.format(a))
			
			if funcao == 'tc':
				tensaoProCalibrador = tc.emf_mVC(float(a), Tref = float(tempAmb))
				assert abs(tensaoProCalibrador) < 150
				print('out {} mV; oper'.format(round(tensaoProCalibrador, 4)))
				cal55XXX.write('out {}E-3V; oper'.format(round(tensaoProCalibrador,4)))
			
			if funcao == 'tempo':
				cal55XXX.write('out {}s; oper'.format(a))

			if funcao == 'livre':
				cal55XXX.write('{}'.format(a))


#########################################################################
# 3458A
#########################################################################

class Swerlein(object):
  """
  PROGRAM MEASURES LOW FREQ RMS VOLTAGES (<1KHz)    01/30/92
  RE-STORE "GODS_AC<RON>"
  
  CAN BE ACCURATE TO 0.001% if Meas_time>30
  NOTES:
  1.  DISTORTED SINEWAVES HAVE HIGHER FREQ HARMONICS
      THAT MAY NOT BE MEASURED if MEASUREMENT BANDWIDTH IS TOO LOW.?
      COMPUTED ERROR IS INCLUDED FOR UP TO 1% HARMONIC DISTORTION.
      THIS ERROR CAN BE REDUCED BY USING SMALLER Aper_targ(LINE 360)
  2.  ESPECIALLY AT LOW SIGNAL LEVELS, TWO AC VOLTMETERS WITH "PERFECT"
      ACCURACY MAY READ DIFFERENTLY if THEY HAVE DIFFERENT
      NOISE BANDWIDTHS.  THIS IS TRUE ONLY if THE SIGNAL BEING MEASURED
      CONTAINS APPRECIABLE HIGH FREQUENCY NOISE OR SPURIOUS SIGNALS.
      THIS PROGRAM DISPLAYS THE MEASUREMENT BANDWIDTH WHEN RUN.
      THIS BANDWIDTH VARIES DEPENDING ON Freq.  IT IS APPROX= .5/Aper_targ.
      ALSO, MAKING Nharm(LINE 370) HIGHER CAN INCREASE BANDWIDTH, BUT
      CAN HURT BASIC ACCURACY BY FORCING SMALL A/D APERTURES.
  
  """
  def __init__(self, Vmeter):
    super(Swerlein, self).__init__()
    self.Vmeter = Vmeter
    self.Vmeter.read_termination = '\n'
    self.Vmeter.timeout = 5000

    self.verbose = False

    self.Meas_time=30.
    self.Forcefreq=False           #1=INPUT FREQ. OF SIGNAL, 0= AUTOMATIC IF>1Hz
    self.Force=False               #1=FORCE SAMP. PARAMETERS, 0= AUTOMATIC
    self.Tsampforce=.001           #  FORCED PARAMETER
    self.Numforce=800.0            #  FORCED PARAMETER
    self.Aper_targ=.001            #A/D APERTURE TARGET (SEC)
    self.Nharm=6                   #MINIMUM # HARMONICS SAMPLED BEFORE ALIAS
    self.Nbursts=30                 #NUMBER OF BURSTS USED FOR EACH Measurement #6
    self.Aperforce=self.Tsampforce-3.E-5 #  FORCED PARAMETER

    self.Freq = None

  @staticmethod
  def Stat(Num,Vmeter):
    Vmeter.write("MEM FIFO;MFORMAT DINT;TARM SGL")
    Vmeter.write("MMATH STAT")
    Vmeter.write("RMATH SDEV")
    Sdev = float(Vmeter.read())
    Vmeter.write("RMATH MEAN")
    Mean = float(Vmeter.read())
    Sdev=math.sqrt(Sdev*Sdev*(Num-1)/Num)     #CORRECT SDEV FORMULA
    return Mean, Sdev

  @staticmethod
  def FNFreq(Expect,Vmeter):
    Vmeter.write("TARM HOLD;LFILTER ON;LEVEL 0,DC;FSOURCE ACDCV;")
    Vmeter.write("FREQ "+str(Expect*.9)+";")
    Vmeter.write("CAL? 245;")
    Cal = float(Vmeter.read())             #FREQUENCY CAL VALUE
    Vmeter.write("TARM SGL;")
    Freq = float(Vmeter.read())
    Freq=Freq/Cal                 #UNCALIBRATED FREQUENCY IS USED
                                  #FOR MORE ACCURATE SAMPLE DETERMINATION
                                  #SINCE TIMER IS UNCALIBRATED
    if Freq==0:
      ##BEEP 2500,.1
      Freq = float(input("FREQ MEASUREMENT WAS 0, PLEASE ENTER THE FREQ:"))
      if Freq>1:
        ##BEEP 200,.1
        print("******************** WARNING# ******************")
        print("WARNING# AUTOMATIC FREQUENCY MEASUREMENT SHOULD HAVE WORKED")
        print("NOTE THAT LEVEL TRIGGERING MAY FAIL")
        print("*************************************************")
    
    return Freq

  @staticmethod
  def FNVmeter_bw(Freq,Range):
    Lvfilter=120000                #LOW VOLTAGE INPUT FILTER B.W.
    Hvattn=36000                   #HIGH VOLTAGE ATTENUATOR B.W.(NUMERATOR)
    Gain100bw=82000                #AMP GAIN 100 B.W. PEAKING CORRECTION#
    if Range<=.12:
      Bw_corr=(1+math.pow((Freq/Lvfilter),2))/(1+math.pow((Freq/Gain100bw),2))
      Bw_corr=math.sqrt(Bw_corr)
    
    if Range>.12 and Range<=12:
      Bw_corr=(1+math.pow((Freq/Lvfilter),2))
      Bw_corr=math.sqrt(Bw_corr)
    
    if Range>12:
      Bw_corr=(1+math.pow((Freq/Hvattn),2))
      Bw_corr=math.sqrt(Bw_corr)
    
    return Bw_corr

  @staticmethod
  def Err_est(self, Freq,Range,Num,Aper,Nbursts):
    #Base IS THE BASIC NPLC 100 DCV 1YR ACCURACY, THIS NUMBER CAN BE
    #SUBSTANTIALLY LOWER FOR HIGH STABILITY OPTION and 90DAY CAL CYCLES
    if Range>120:                      #SELF HEATING +BASE ERROR
      Base=15.0
    else:
      Base=10.0                        #BASIC 1YR ERROR(ppm)
    

    #Vmeter_bw IS ERROR DUE TO UNCERTAINTY IN KNOWING THE HIGH FREQUENCY
    #RESPONSE OF THE DCV FUNCTION FOR VARIOUS RANGES and FREQUENCIES
    #UNCERTAINTY IS 30% and THIS ERROR IS RANDOM
    X1=self.FNVmeter_bw(Freq,Range)
    X2=self.FNVmeter_bw(Freq*1.3,Range)
    Vmeter_bw=int(1.E+6*abs(X2-X1))  #ERROR DUE TO METER B.W.

    #Aper_er IS THE DCV GAIN ERROR FOR VARIOUS A/D APERTURES
    #THIS ERROR IS SPECIFIED IN A GRAPH ON PAGE 11 OF THE DATA SHEET
    #THIS ERROR IS RANDOM
    Aper_er=int(1.0*.002/Aper)       #GAIN UNCERTAINTY - SMALL A/D APERTURE
    if Aper_er>30 and Aper>=1.E-5:
      Aper_er=30.0
    
    if Aper<1.E-5:
      Aper_er=10.0+int(.0002/Aper)
    
    #Sincerr IS THE ERROR DUE TO THE APERTURE TIME NOT BEING PERFECTLY KNOWN
    #THIS VARIATION MEANS THAT THE Sinc CORRECTION TO THE SIGNAL FREQUENCY
    #IS NOT PERFECT.  ERROR COMPONENTS ARE CLOCK FREQ UNCERTAINTY(0.01%)
    #AND SWITCHING TIMING (50ns).  THIS ERROR IS RANDOM.
    X=math.pi*Aper*Freq
    Sinc=math.sin(X)/X
    Y=math.pi*Freq*(Aper*1.0001+5.0E-8)
    Sinc2=math.sin(Y)/Y
    Sincerr=int(1.E+6*abs(Sinc2-Sinc))    #APERTURE UNCERTAINTY ERROR

    #Dist IS ERROR DUE TO UP TO 1% DISTORTION OF THE INPUT WAVEFORM
    #if THE INPUT WAVEFORM HAS 1% DISTORTION, THE ASSUMPTION IS MADE
    #THAT THIS ENERGY IS IN THE THIRD HARMONIC.  THE APERTURE CORRECTION,
    #WHICH IS MADE ONLY FOR THE FUNDAMENTAL FREQUENCY WILL THEN BE
    #INCORRECT.  THIS ERROR IS RETURNED SEPERATELY.
    X=math.pi*Aper*Freq
    Sinc=math.sin(X)/X
    Sinc2=math.sin(3.0*X)/3.0/X      #SINC CORRECTION NEEDED FOR 3rd HARMONIC
    Harm_er=abs(Sinc2-Sinc)
    Dist=math.sqrt(1.0+math.pow((.01*(1+Harm_er)),2.0))-math.sqrt(1+math.pow(.01,2.0))
    Dist=int(Dist*1.E+6)

    #Tim_er IS ERROR DUE TO MISTIMING.  IT CAN BE SHOWN THAT if A
    #BURST OF Num SAMPLES ARE USED TO COMPUTE THE RMS VALUE OF A SINEWAVE
    #AND THE SIZE OF THIS BURST IS WITHIN 50ns*Num OF AN INTEGRAL NUMBER
    #OF PERIODS OF THE SIGNAL BEING MEASURED, AN ERROR IS CREATED
    #BOUNDED BY 100ns/4/Tsamp.  THIS ERROR IS DUE TO THE 100ns QUANTIZATION
    #LIMITATION OF THE HP3458A TIME BASE.  if THIS ERROR WERE ZERO,:
    #Num*Tsamp= INTEGER/Freq, BUT WITH THIS ERROR UP TO 50ns OF TIMEBASE
    #ERROR IS PRESENT PER SAMPLE, THEREFORE TOTAL TIME ERROR=50ns*Num
    #THIS ERROR CAN ONLY ACCUMULATE UP TO 1/2 *Tsamp, AT WHICH POINT THE
    #ERROR IS BOUNDED BY 1/4/Num
    #THIS ERROR IS FURTHER REDUCED BY USING THE LEVEL TRIGGER
    #TO SPACE Nbursts AT TIME INCREMENTS OF 1/Nbursts/Freq.  THIS
    #REDUCTION IS SHOWN AS 20*Nbursts BUT IN FACT IS USUALLY MUCH BETTER
    #THIS ERROR IS ADDED ABSOLUTELY TO THE Err CALCULATION
    Tim_er=int(1.E+6*1.E-7/4/(Aper+3.E-5)/20.0)#ERROR DUE TO HALF CYCLE ERROR
    Limit=int(1.E+6/4.0/Num/20.0)
    if Tim_er>Limit:
      Tim_er=Limit

    #Noise IS THE MEASUREMENT TO MEASUREMENT VARIATIONS DUE TO THE
    #INDIVIDUAL SAMPLES HAVING NOISE.  THIS NOISE IS UNCORRELATED AND
    #IS THEREFORE REDUCED BY THE SQUARE ROOT OF THE NUMBER OF SAMPLES
    #THERE ARE Nbursts*Num SAMPLES IN A MEASUREMENT.  THE SAMPLE NOISE IS
    #SPECIFIED IN THE GRAPH ON PAGE 11 OF THE DATA SHEET.  THIS GRAPH
    #SHOWS 1 SIGMA VALUES, 2 SIGMA VALUES ARE COMPUTED BELOW.
    #THE ERROR ON PAGE 11 IS EXPRESSED AS A % OF RANGE and IS MULTIPLIED
    #BY 10 SO THAT IT CAN BE USED AS % RDG AT 1/10 SCALE.
    #ERROR IS ADDED IN AN ABSOLUTE FASHION TO THE Err CALCULATION SINCE
    #IT WILL APPEAR EVENTUALLY if A MEASUREMENT IS REPEATED OVER and OVER
    Noiseraw=.9*math.sqrt(.001/Aper)       #1 SIGMA NOISE AS PPM OF RANGE
    Noise=Noiseraw/math.sqrt(Nbursts*Num)  #REDUCTION DUE TO MANY SAMPLES
    Noise=10.0*Noise                   #NOISE AT 1/10 FULL SCALE
    Noise=2.0*Noise                    #2 SIGMA
    if Range<=.12:               #NOISE IS GREATER ON 0.1 V RANGE
      Noise=7.0*Noise                  #DATA SHEET SAYS USE 20, BUT FOR SMALL
      Noiseraw=7.0*Noiseraw            #APERTURES, 7 IS A BETTER NUMBER
    
    Noise=int(Noise)+2.0                   #ERROR DUE TO SAMPLE NOISE

    #Df_err IS THE ERROR DUE TO THE DISSIPATION FACTOR OF THE P.C. BOARD
    #CAPACITANCE LOADING DOWN THE INPUT RESISTANCE.  THE INPUT RESISTANCE
    #IS 10K OHM FOR THE LOW VOLTAGE RANGES and 100K OHM FOR THE HIGH VOLTAGE
    #RANGES (THE 10M OHM INPUT ATTENUATOR).  THIS CAPACITANCE HAS A VALUE
    #OF ABOUT 15pF and A D.F. OF ABOUT 1.0%.  IT IS SWAMPED BY 120pF
    #OF LOW D.F. CAPACITANCE (POLYPROPALENE CAPACITORS) ON THE
    #LOW VOLTAGE RANGES WHICH MAKES FOR AN EFFECTIVE D.F. OF ABOUT .11%.
    #THIS CAPACITANCE IS SWAMPED BY 30pF LOW D.F. CAPACITANCE ON THE
    #HIGH VOLTAGE RANGES WHICH MAKES FOR AN EFFECTIVE D.F. OF .33%.
    #THIS ERROR IS ALWAYS IN THE NEGATIVE DIRECTION, SO IS ADDED ABSOLUTELY
    if Range<=12:
      Rsource=10000.0
      Cload=1.33E-10
      Df=1.1E-3          #0.11%
    else:
      Rsource=1.E+5
      Cload=5.0E-11
      Df=3.3E-3          #0.33%
    
    Df_err=2.0*math.pi*Rsource*Cload*Df*Freq
    Df_err=int(1.E+6*Df_err)#ERROR DUE TO TO PC BOARD DIELECTRIC ABSORBTION

    #Err IS TOTAL ERROR ESTIMATION.  RANDOM ERRORS ARE ADDED IN RSS FASHION
    Err=math.sqrt(math.pow(Base,2)+math.pow(Vmeter_bw,2)+math.pow(Aper_er,2)+math.pow(Sincerr,2))
    Err=int(Err+Df_err+Tim_er+Noise)            #TOTAL ERROR (ppm)

    return Err, Dist

  @staticmethod
  def Samp_parm(Freq,Meas_time,Aper_targ,Nharm_min,Nbursts):
    Aper=Aper_targ
    Tsamp=1.E-7*int((Aper+3.0E-5)/1.E-7+.5)       #ROUND TO 100ns
    Submeas_time=Meas_time/Nbursts                #TARGET TIME PER BURST
    Burst_time=Submeas_time*Tsamp/(.0015+Tsamp)   #IT TAKES 1.5ms FOR EACH
    #                                              SAMPLE TO COMPUTE Sdev
    Approxnum=int(Burst_time/Tsamp+.5)
    Ncycle=int(Burst_time*Freq+.5)                ## OF 1/Freq TO SAMPLE
    if Ncycle==0:
      Ncycle=1
      Tsamp=1.E-7*int(1.0/Freq/Approxnum/1.E-7+.5)  #TIME BETWEEN SAMPLES
      Nharm=int(1./Tsamp/2./Freq)     ## HARMONICS BEFORE ALIAS OCCURS
      if Nharm<Nharm_min:       #NEED TO INCREASE SAMPLE FREQUENCY
        Nharm=Nharm_min
        Tsamp=1.E-7*int(1./2./Nharm/Freq/1.E-7+.5)
      
    else:
      Nharm=int(1./Tsamp/2./Freq)     ## HARMONICS BEFORE ALIAS OCCURS
      if Nharm<Nharm_min:       #NEED TO INCREASE SAMPLE FREQUENCY
        Nharm=Nharm_min
      
      Tsamptemp=1.E-7*int(1./2./Nharm/Freq/1.E-7+.5)  #FORCE ALIAS TO OCCUR

      Burst_time=Submeas_time*Tsamptemp/(.0015+Tsamptemp)
      Ncycle=int(Burst_time*Freq+1.)
      Num=int(Ncycle/Freq/Tsamptemp+.5)
      if Ncycle>1:
        K=int(Num/20./Nharm+1)
      else:
        K=0
      
      Tsamp=1.E-7*int(Ncycle/Freq/(Num-K)/1.E-7+.5)  #NOW ALIAS OCCURS
      #                                     MUCH HIGHER THAN Nharm*Freq
      #                                     K WAS PICKED TO TRY and MAKE
      #                                     ALIAS ABOUT 10*Nharm*Freq
      if Tsamp-Tsamptemp<1.E-7:
        Tsamp=Tsamp+1.E-7
    
    Aper=Tsamp-3.E-5
    Num=int(Ncycle/Freq/Tsamp+.5)
    if Aper>1:
      Aper=1.0       #MAX APERTURE OF HP3458A
    
    try:
      assert Aper>=1.E-6          #MIN APERTURE OF HP3458A
    except Exception as e:
      print("************** ERROR *****************")
      print("A/D APERTURE IS TOO SMALL")
      print("LOWER Aper_targ, Nharm, OR INPUT Freq")
      print("***************************************")
      raise e
    
    return Tsamp, Aper, Num

  def new_signal(self, Expect):
    # limita casas decimais
    Expect = round(Expect, 3)
    assert Expect >= 0.1e-3
    assert Expect <= 700.0
    ##OUTPUT 2 USING "#,B,B";255,75       #CLEAR DISPLAY
    self.Expect=1.55*Expect           #PEAK VALUE *1.1
    if self.Expect<.12:
      self.Range=.1

    if self.Expect>=.12 and self.Expect<1.2:
      self.Range=1.0

    if self.Expect>=1.2 and self.Expect<12:
      self.Range=10.0

    if self.Expect>=12 and self.Expect<120:
      self.Range=100.0

    if self.Expect>=120:
      self.Range=1000.0

    ##OUTPUT 2 USING "#,B,B";255,75        #CLEAR DISPLAY
    self.Vmeter.clear()
    self.Vmeter.write("RESET;DCV 1000")
    self.Freq = None

  def read(self):
    ##CLEAR Vmeter
    #self.Vmeter.write("DISP OFF,***GOD'S_AC***")
    self.Vmeter.write("DISP OFF,***LME'S_AC***")

    if self.Freq == None:
      self.Vmeter.write("RESET;")
      time.sleep(0.2)
      if self.Forcefreq:            #if MANUALLY ENTERING FREQUENCY
        self.Freq = float(input("ENTER FREQ:"))
      else:
        self.Freq=self.FNFreq(self.Expect,self.Vmeter)     #GET INPUT SIGNAL FREQUENCY

    if self.Force:                  #if NOT GETTING PARAMETERS AUTOMATICALLY
      Tsamp=self.Tsampforce
      Aper=self.Aperforce
      Num=self.Numforce
    else:                             #AUTOMATICALLY GET SAMPLING PARAMETERS
      Tsamp, Aper, Num = self.Samp_parm(self.Freq,self.Meas_time,self.Aper_targ,self.Nharm,self.Nbursts)
    #
    # SETUP HP3458A
    #
    self.Vmeter.write("TARM HOLD;AZERO OFF;DCV "+str(self.Range))
    self.Vmeter.write("APER "+str(Aper)+";NRDGS "+str(Num)+",TIMER")
    self.Vmeter.write("TIMER "+str(Tsamp))
    self.Vmeter.write("TRIG LEVEL;LEVEL 0,DC;DELAY 0;LFILTER ON")
    self.Vmeter.write("MSIZE?")
    Storage = float(self.Vmeter.read().split(',')[0])
    Storage=int(Storage/4.)     #STORAGE CAPACITY IN VOLTMETER (DINT DATA)
    try:
      assert Num<=Storage
    except Exception as e:
      print("******** NOT ENOUGH VOLTMETER MEMORY FOR NEEDED SAMPLES ***")
      print("         TRY A LARGER Aper_targ VALUE OR SMALLER Num")
      raise e

    time.sleep(0.1)

    #
    # PRELIMINARY COMPUTATIONS
    #

    X=math.pi*Aper*self.Freq
    Sinc=math.sin(X)/X                     #USED TO CORRECT FOR A/D APERTURE ERROR
    Bw_corr=self.FNVmeter_bw(self.Freq,self.Range)   #USED TO CORRECT FOR Vmeter BANDWIDTH
    Err, Dist_er = self.Err_est(self,self.Freq,self.Range,Num,Aper,self.Nbursts)#MEASUREMENT UNCERTAINTY
    
    # if self.Force and self.verbose:
      # print("****** PARAMETERS ARE FORCED, ACCURACY MAY BE DEGRADED# ******")

    if self.verbose:
      # print("SIGNAL FREQUENCY(Hz)= "+str(self.Freq))
      # print("Number of samples in each of"+str(self.Nbursts)+";bursts= "+str(Num))
      # print("Sample spacing(sec)= "+str(Tsamp))
      # print("A/D Aperture(sec)= "+str(Aper))
      # print("Measurement bandwidth(Hz)= "+str(int(5.0/Aper)/10.0))
      # print("ESTIMATED TOTAL SINEWAVE MEASUREMENT UNCERTAINTY(ppm)= "+str(Err))
      print("ADDITIONAL ERROR FOR 1% DISTORTION(3rd HARMONIC)(ppm)= "+str(Dist_er))
      print("NOTE: ERROR ESTIMATE ASSUMES (ACAL DCV) PERFORMED RECENTLY(24HRS)")
    #----------------------------------------
    #self.Vmeter.write("DISP OFF,***GOD'S_AC***")
    #self.Vmeter.write("DISP OFF,***LME'S_AC***")
    
    # frases = ["AI_AI", "PAPAGAIADA", "QUE_CALORZINHO", "SO_LAGRIMAS", "E_FOGO", "JOOOOVEM!", "QUE?", "XOOOVEM", "GALERINHA!"]
    # random.shuffle(frases)
    # self.Vmeter.write("DISP OFF,"+frases[0][0:14])
    
    self.Vmeter.write("DISP OFF, swerleinISTA")

    
    if self.verbose:
      print("The "+str(self.Nbursts)+" intermediate results:")
    
    Begin=time.time()
    Sum=0
    Sumsq=0
    for I in range(0, self.Nbursts):
      Delay=float(I)/float(self.Nbursts)/self.Freq+1.E-6 # NOTE: Very Important the float conversion 
      #                                                    when using python2, so numbers dont get truncated
      self.Vmeter.write("DELAY "+str(Delay))
      self.Vmeter.write_raw("TIMER "+str(Tsamp))
      Mean, Sdev = self.Stat(Num,self.Vmeter)        #MAKE MEASUREMENT
      Sumsq=Sumsq+Sdev*Sdev+Mean*Mean
      Sum=Sum+Mean
      Temp=Sdev*Bw_corr/Sinc               #CORRECT A/D Aper and Vmeter B.W.
      Temp=self.Range/1.E+7*int(Temp*1.E+7/self.Range)#6 DIGIT TRUNCATION
      
      if self.verbose:
        print(str(Temp).replace('.',','))

    Dcrms=math.sqrt(Sumsq/float(self.Nbursts))
    Dc=Sum/self.Nbursts
    Acrms=Dcrms*Dcrms-Dc*Dc
    if Acrms<0:
      Acrms=0                #PROTECTION FOR SQR OF NEG NUMBER

    Acrms=math.sqrt(Acrms)
    Acrms=Acrms*Bw_corr/Sinc               #CORRECT A/D Aper and Vmeter B.W.
    Dcrms=math.sqrt(Acrms*Acrms+Dc*Dc)
    End=time.time()
    ##Acrms=self.Range/1.E+7*int(Acrms*1.E+7/self.Range+.5)    #6 DIGIT TRUNCATION
    ##Dcrms=self.Range/1.E+7*int(Dcrms*1.E+7/self.Range+.5)    #6 DIGIT TRUNCATION

    #********************** print RMS VALUES ***************
    if self.verbose:
      print("AC RMS VOLTAGE= "+str(Acrms))
      # print("ACDC RMS VOLTAGE= "+str(Dcrms))
      # print("MEASUREMENT TIME(sec)= "+str(End-Begin))

    self.Vmeter.write("DISP OFF,"+"'"+str(Acrms)+" VAC'")
    
    res = dict({
      'Acrms': Acrms,
      'Dcrms': Dcrms,
      'time':End-Begin,
      'Freq': self.Freq,
      'sinewave_uncertainty_ppm': Err,
      'distortion_error_ppm': Dist_er
    })

    return res

def leitura3458A(endGpib, endereco):

	rm = visa.ResourceManager()
	dmm3458A = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')
	
	print('Qual função?')
	print('vdc | vac | idc | iac | res4f | Math null | freq | swerlein')
	funcao = input('> ').lower()

	if funcao == 'voltar':
		main()

	if funcao[0:1] == 'm':
		mathNull3458A(dmm3458A)
		leitura3458A(endGpib, endereco)

	if funcao[0:3] == 'swe':
		funcao = 'swerlein'
	# 	leituraSwerlein(dmm3458A)
	# 	leitura3458A(endGpib, endereco)
	
	if funcao[0:1] == 'r':
		funcao = 'res4f'
	
	configuracoes = {
		'vdc': 'DCV -1 ; NPLC 100; DELAY 2;TARM HOLD;',
		'vac': 'ACV -1 ; NPLC 100; SETACV SYNC; LFILTER ON; DELAY 2;TARM HOLD;',
		'idc': 'DCI -1 ; NPLC 100; DELAY 2;TARM HOLD;',
		'iac': 'ACI -1 ; NPLC 100; DELAY 2;TARM HOLD;',
		'res4f': 'OHMF -1 ; NPLC 100; OCOMP ON; DELAY 2;TARM HOLD;',
		'freq': 'FREQ -1 .00001; TARM HOLD;',
		'swerlein': 'DCV -1 ; NPLC 100; DELAY 2;TARM HOLD;',
	}
	
	configuracao = configuracoes[funcao]

	# de tensão dc tem tb
	if funcao == 'vdc':
		print('\n## Medir tensão DC com 3458A ##')
		print('\n## Por acaso você está \nmedindo um simulador de pH\n que precisa de alta \nimpedância? (s/n)')
		depH = input('> ')

		if depH == 's':
			configuracao = "DCV -1 ; NPLC 100; DELAY 2;TARM HOLD; FIXEDZ OFF"

	if funcao == 'freq':
		dmm3458A.write(configuracoes['vac'])


	print('\n> Multímetro configurado como:')
	dmm3458A.write('{}'.format(configuracao))
	print(configuracao)

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'

	while True:
		
		if mesmoSetup == 'n':
		
			if funcao[0:3] == 'swe':
				newExpect = input("> VRMS esperado ± 10 %:\n> ").replace(',','.')
				if newExpect.find('m') > -1:
					newExpect = str(float(newExpect.split('m')[0])/1_000)
					print(newExpect)

			setup = defineSetup(nplc=True, ajustaCasas=True)
			nplc = setup['nplc']
			ajustaCasas = setup['ajustaCasas']
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']
		
		# if funcao[0:3] == 'swe':
		# 	input("Aplique o sinal, espere estabilizar e pressione 'Enter'")
		# 	print('\n\n')

		print('\n[3458A] '+quantasLeituras+'≡ '+unidadeSelecionada+' nplc '+nplc)
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		dmm3458A.write('{}'.format(configuracao))
		dmm3458A.write('NPLC {}'.format(nplc))

		if funcao[0:3] == 'swe':
			swerlein = Swerlein(dmm3458A)
			swerlein.verbose = False
			#swerlein.Meas_time=30.
			Expect = None
				
			for leitura in range(int(quantasLeituras)):
			
				if newExpect != "":
					Expect = newExpect
					assert Expect != None
					swerlein.new_signal(float(Expect))
				
				res = swerlein.read()

				a = res['Acrms']
				processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)

		else:

			for leitura in range(int(quantasLeituras)):
				a = float(dmm3458A.query('TARM HOLD;TARM SGL;'))
				processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
			
		proxPto = proximoPonto(leiturasPonto_txt, leituras_txt, dmm3458A, endGpib, temReset=True)
		
		mesmoSetup = proxPto['mesmoSetup']
		resetFuncao = proxPto['resetFuncao']

		if resetFuncao == True:
			leitura3458A(endGpib, endereco)


def mathNull3458A(dmm3458A):

	print('\n## Math Null 3458A ##')

	print('Qual função?')
	print('vdc | vac | idc | iac | res4f')
	funcao = input('> ').lower()

	if funcao[0:1] == 'r':
		funcao = 'res4f'
	
	configuracoes = {
		'vdc': 'DCV -1 ; NPLC 100; DELAY 2;TARM HOLD;',
		'vac': 'ACV -1 ; NPLC 100; SETACV SYNC; LFILTER ON; DELAY 2;TARM HOLD;',
		'idc': 'DCI -1 ; NPLC 100; DELAY 2;TARM HOLD;',
		'iac': 'ACI -1 ; NPLC 100; DELAY 2;TARM HOLD;',
		'res4f': 'OHMF -1 ; NPLC 100; OCOMP ON; DELAY 2;TARM HOLD;',
	}
	
	configuracao = configuracoes[funcao]

	print('\n> Multímetro configurado como:')
	dmm3458A.write('{}'.format(configuracao))
	print(configuracao)
	ajustaCasas = 1 # a escala esperada é a menor possível e esse parâmetro não afeta a ação do multímetro

	input('\nPressione ENTER para iniciar')
	
	executar = True

	while executar == True:
	
		print('\n> 1 de 3\n>> Coletando pré Math Null')

		a = float(dmm3458A.query('TARM HOLD;TARM SGL;'))
		a = str(round(a/float(ajustaCasas),9))
		a = a.replace('.', ',')
		print('> pré:')
		print(a)
	
		dmm3458A.write('MATH NULL;')
		print('\n> 2 de 3\n>> Math null realizado')
		print('\n> 3 de 3\n>> Coletando pós Math Null')

		a = float(dmm3458A.query('TARM HOLD;TARM SGL;'))
		a = str(round(a/float(ajustaCasas),9))
		a = a.replace('.', ',')
		print('> pós:')
		print(a)

		a = input('\nRepetir Math null? (s/n)\n> ').lower()

		if a == 'n':
			executar = False

#########################################################################
# 5335A
#########################################################################

def leitura3553A(endGpib, endereco):

	rm = visa.ResourceManager()
	cont5335A = rm.open_resource("GPIB{}::{}::INSTR".format(endGpib, endereco), read_termination = "\r\n", timeout = 20000, write_termination = "\r\n")

	input('\n## Configure NO INSTRUMENTO o gate e o trigger e pressione ENTER ##')

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'
	
	while True:

		if mesmoSetup == 'n':
		
			setup = defineSetup(ajustaCasas=True)
			ajustaCasas = setup['ajustaCasas']
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']	

			print('\n[3553A] '+quantasLeituras+'≡ '+unidadeSelecionada)
			leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
			setPonto = []

			cont5335A.write("RE")
			time.sleep(7) # tempo do reset

			for leitura in range(int(quantasLeituras)):

				a = float(cont5335A.read().split(" ")[-1])
				processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
			
			mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, cont5335A, endGpib)
	

#########################################################################
# PM6304
#########################################################################

def livrePM6304(endGpib, endereco):
	
	rm = visa.ResourceManager()
	pntPM6304 = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n', write_termination='\n')

	print('\nPonte configurada como AUTO, DC BIAS OFF e AVG ON')

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'

	while True:
		
		if mesmoSetup == 'n':

			print('\nQual a grandeza a ser medida?')
			print('1 - Capacitância')
			print('2 - Resistência')
			print('3 - Indutância')
			grandeza = input('> ')
			

			if grandeza == '1':
				grandeza = 'cap?'
				grandezaPrint = 'Farad'
			
			if grandeza == '2':
				grandeza = 'resi?'
				grandezaPrint = 'Ohm'

			if grandeza == '3':
				grandeza = 'indu?'
				grandezaPrint = 'Henry'

			print('\nQual a frequência em Hz?')
			frequencia = input('> ')

			setup = defineSetup()
			quantasLeituras = setup['quantasLeituras']

			print('\n> Tempo entre leituras? (segundos)')
			tempoEspera = input('> ')

		pntPM6304.write('AUTO')
		time.sleep(2)
		pntPM6304.write('DC BIAS OFF')
		time.sleep(2)
		pntPM6304.write('AVG ON')
		time.sleep(2)
		pntPM6304.write('CON')
		time.sleep(2)
		pntPM6304.write('FRE {}'.format(frequencia))
		time.sleep(2)

		print('\n[PM6304] '+quantasLeituras+'≡ '+grandezaPrint+' ('+frequencia+' Hz - '+tempoEspera+' s)')
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		for leitura in range(int(quantasLeituras)):
			
			a = pntPM6304.query('{}'.format(grandeza))
			time.sleep(int(tempoEspera))
		
			if a.find('E') < 0:
				a = a + 'E '	

			unidades = {
				'-12': 'p',
				'-9': 'n',
				'-6': 'µ',
				'-3': 'm',
				' ': '',
				'3': 'k',
				'6': 'M',
				'9': 'G'
			}
			unidadeLeitura = a.split('E')[1]
			unidadeLeitura = unidades[unidadeLeitura]
			
			a = a.split('E')[0]
			a = a.split(' ')[1]
			a = a.replace('.', ',')

			indice = str(leitura+1)
			indice = '[{}/{}]'.format(indice, quantasLeituras)
			grandezaPrint = grandezaPrint[0] if grandezaPrint[0] != 'O' else grandezaPrint[0:3]
			print('{}\t{} {}{}'.format(indice, a, unidadeLeitura, grandezaPrint))
			leiturasPonto_txt.write(a+'\n')
			leituras_txt.write(a+'\n')
		
		
		pntPM6304.write('SIN')
		time.sleep(2)

		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, pntPM6304, endGpib)
	

#########################################################################
# 8508A
#########################################################################

def leitura8508A(endGpib, endereco):

	rm = visa.ResourceManager()
	dmm8508A = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')

	print('\n## Medir com 8508A ##')

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'
	
	while True:
		
		if mesmoSetup == 'n':
			setup = defineSetup(ajustaCasas=True)
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']
			ajustaCasas = setup['ajustaCasas']

		print('\n[8508A] '+quantasLeituras+'≡ '+unidadeSelecionada)
		
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		
		setPonto = []
		for leitura in range(int(quantasLeituras)):
			a = float(dmm8508A.query("TRG_SRCE EXT;X?"))
			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, dmm8508A, endGpib)

#########################################################################
# 34410A
#########################################################################

def leitura344XXX(endGpib, endereco):

	rm = visa.ResourceManager()
	dmm344xxx = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')

	input('\n> Configure o multímetro e pressione ENTER')
	
	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	# leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')
	leituras_txt = open('I:/LME/TEMPORARIO/Diogo/Automacao/controlista/leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')


	mesmoSetup = 'n'

	while True:
		
		if mesmoSetup == 'n':
		
			setup = defineSetup(ajustaCasas=True)
			ajustaCasas = setup['ajustaCasas']
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']

		print('[344XXX] '+quantasLeituras+'≡ '+unidadeSelecionada)
		# leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		leiturasPonto_txt = open('I:/LME/TEMPORARIO/Diogo/Automacao/controlista/leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		dmm344xxx.write('TRIG:SOUR IMM') ## seta para trigger interno

		for leitura in range(int(quantasLeituras)):
			a = float(dmm344xxx.query('read?'))
			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		dmm344xxx.write('TRIG:SOUR EXT') ## seta para trigger externo

		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, dmm344xxx, endGpib)

#########################################################################
# 88XXX
#########################################################################

def leitura88XXX(endGpib, endereco):
	
	rm = visa.ResourceManager()
	dmm88XXX = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')
	dmm88XXX.write('SYST:ERR:BEEP OFF')

	input('\n> Configure o multímetro e pressione ENTER')

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')
	
	mesmoSetup = 'n'

	while True:

		if mesmoSetup == 'n':
			setup = defineSetup(ajustaCasas=True)
			ajustaCasas = setup['ajustaCasas']
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']

		print('\n[88XX] '+quantasLeituras+'≡ '+unidadeSelecionada)
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		dmm88XXX.write('TRIG:SOUR IMM') ## seta para trigger interno

		for leitura in range(int(quantasLeituras)):
			
			a = float(dmm88XXX.query('read?'))
			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		dmm88XXX.write('TRIG:SOUR EXT') ## seta para trigger externo
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, dmm88XXX, endGpib)


#########################################################################
# 7561
#########################################################################

def leitura7561(endGpib, endereco):
	
	rm = visa.ResourceManager()
	dmm7561 = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=20000, read_termination='\n')
	dmm7561.write('SYST:ERR:BEEP OFF')

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'

	while True:

		if mesmoSetup == 'n':
					
			setup = defineSetup()
			quantasLeituras = setup['quantasLeituras']

		ajustaCasas = 1
		
		print('\n[3458A] '+quantasLeituras+'≡ ')
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []
		
		for leitura in range(int(quantasLeituras)):
			a = dmm7561.query('read?')
			a = a.split('+')[1]
			a = a.split('E')[0]
			a = float(a)
			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, dmm7561, endGpib)

#########################################################################
# 5790A
#########################################################################

def leitura5790A(endGpib, endereco):

	rm = visa.ResourceManager()
	dmm5790A = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')
		
	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'
	
	while True:
		
		if mesmoSetup == 'n':

			print('\nUsar qual digital mode?')
			print('(leitura = média de x leituras)')
			print('off | slow | medium | fast')
			digitalMode = input('> ')
			if digitalMode[0] == 'o':
				digitalMode = 'off'

			if digitalMode[0] == 's':
				digitalMode = 'slow' 

			if digitalMode[0] == 'm':
					digitalMode = 'medium'

			if digitalMode[0] == 'f':
				digitalMode = 'fast'

			print('\nUsar qual filter restart?')
			print('(avalia variação do conjunto de\nleituras antes do calculo da média)')
			print('coarse | medium | fine')
			filterRestart = input('> ')

			if filterRestart[0] == 'c':
				filterRestart = 'coarse'

			if filterRestart[0] == 'm':
				filterRestart = 'medium'
			
			if filterRestart[0] == 'f':
				filterRestart = 'fine' 

			
			setup = defineSetup(ajustaCasas=True)
			ajustaCasas = setup['ajustaCasas']
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']
			
		print('\n[5790A] '+quantasLeituras+'≡ '+unidadeSelecionada + ' ' + digitalMode + ' ' + filterRestart)
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		for leitura in range(int(quantasLeituras)):

			dmm5790A.write('hires on')
			dmm5790A.write('DFILT {}, {}'.format(digitalMode, filterRestart))
			
			if leitura == 0:
				dmm5790A.write('RANGE {}V;'.format(unidadeSelecionada[0]))
				a = dmm5790A.query("TRIG; *WAI; VAL?")
				a = float(a.split(',')[0])

				# if a > 700e-3:
				# 	ajustaCasas = 1
				# else:
				# 	ajustaCasas = 1/1000
				
			a = dmm5790A.query("TRIG; *WAI; VAL?")
			a = float(a.split(',')[0])
			time.sleep(1)

			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, dmm5790A, endGpib)

#########################################################################
# 34420A
#########################################################################

def leitura34420A(endGpib, endereco):

	rm = visa.ResourceManager()
	dmm34420A = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=20000, read_termination='\n')

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'
	
	while True:
		
		if mesmoSetup == 'n':
				
			setup = defineSetup(ajustaCasas=True)
			ajustaCasas = setup['ajustaCasas']
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']
		
				
		print('\n[3458A] '+quantasLeituras+'≡ '+unidadeSelecionada)
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		for leitura in range(int(quantasLeituras)):
			
			a = float(dmm34420A.query('read?'))
			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, dmm34420A, endGpib)


#########################################################################
# QuadTech 1920
#########################################################################

def leituraQuadtech(endGpib, endereco):

	rm = visa.ResourceManager()
	quadTech = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')
		
	print('\n## Leituras com QuadTech 1920 ##')

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'
	
	while True:
		
		if mesmoSetup == 'n':
				
			print('\n> Primeira ou Segunda medição?')
			qualMedicao = input('> ').lower()

			if qualMedicao == '1' or qualMedicao[0] == 'p':
				qualMedicaoPrint = 'Medida 1'
			else:
				qualMedicaoPrint = 'Medida 2'
			
			setup = defineSetup(ajustaCasas=True)
			ajustaCasas = setup['ajustaCasas']
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']
			
		print('\n[QuadTech1920] '+quantasLeituras+'≡ '+unidadeSelecionada+' ('+qualMedicaoPrint+')')
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		for leitura in range(int(quantasLeituras)):

			time.sleep(4)
			quadLeu = quadTech.query('M?')

			if qualMedicao == '1' or qualMedicao[0] == 'p':
				try:
					a = float(quadLeu.split(',')[1])
				except Exception as e:
					a = 0
			
			if qualMedicao != '1' or qualMedicao[0] == 's':
				try:
					a = float(quadLeu.split(',')[2])
				except Exception as e:
					a = 0
			
			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
	
		quadTech.write('STOP')
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, quadTech, endGpib)


#########################################################################
# 34970A
#########################################################################

def leitura34970A():

	rm = visa.ResourceManager()
	#listaInstrumentos()

	endGpib = '0'
	endGpib = input('Qual interface GPIB? \n> ')

	endereco = input('Qual endereço GPIB?\n> ')
	if endereco[0] == 'n':
		listaInstrumentos(endGpib)
		endereco = input('Qual endereço GPIB?\n> ')
	dmm34970A = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=20000, read_termination='\n')
	dmm34970A.write('*CLS')
		
	print('\n## Leituras com 34970A ##')
	print('\n')
	input('Faça a configuração do canal e pressione ENTER')	

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	print(horaData)
	
	if endGpib == '0':
		fileCompleta = open('I:\LME\TEMPORARIO\Diogo\Automacao\controlista\Leituras_{}.txt'.format(horaData), 'w', encoding='utf-8')

	if endGpib != '0':
		fileCompleta = open('I:\LME\TEMPORARIO\Diogo\Automacao\controlista\Leituras1_{}.txt'.format(horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'

	
	while True:
		
		if mesmoSetup == 'n':
				
			print('\n> Qual canal será calibrado?')
			qualCanal = input('> ')

			print('\n> Qual unidade no Autolab?')
			print('> 1 -> nano')
			print('> 2 -> micro')
			print('> 3 -> mili')
			print('> 4 -> sem multiplicador')
			print('> 5 -> kilo')
			print('> 6 -> Mega')
			print('> 7 -> Giga')
			ajustaCasas = input('> ')
		

		
			if ajustaCasas == '1' or ajustaCasas == 'nano' or ajustaCasas == 'n':
				unidadeSelecionada = 'nano'
				ajustaCasas = 1/1000000000
			
			if ajustaCasas == '2' or ajustaCasas == 'micro' or ajustaCasas == 'u':
				unidadeSelecionada = 'micro'
				ajustaCasas = 1/1000000
			
			if ajustaCasas == '3' or ajustaCasas == 'mili' or ajustaCasas == 'm':
				unidadeSelecionada = 'mili'
				ajustaCasas = 1/1000
			
			if ajustaCasas == '4' or ajustaCasas == 'sem' or ajustaCasas == '':
				unidadeSelecionada = 'sem mult.'
				ajustaCasas = 1

			if ajustaCasas == '5' or ajustaCasas == 'kilo' or ajustaCasas == 'k':
				unidadeSelecionada = 'kilo'
				ajustaCasas = 1000
			
			if ajustaCasas == '6' or ajustaCasas == 'mega' or ajustaCasas == 'M':
				unidadeSelecionada = 'mega'
				ajustaCasas = 1000000

			if ajustaCasas == '7' or ajustaCasas == 'giga' or ajustaCasas == 'G':
				unidadeSelecionada = 'giga'
				ajustaCasas = 1000000

		
		# print('\n> Unidade selecionada: '+unidadeSelecionada)				
		# print('> Canal selecionado: {}'.format(qualCanal))
		# print('> Número de leituras: '+quantasLeituras)
		# print('> Range selecionado: {}'.format(unidadeSelecionada))

		print('\n[ch'+qualCanal+'] '+quantasLeituras+'≡ '+unidadeSelecionada)

		dmm34970A.write('*CLS')
		
				
		if endGpib == '0':
			file = open('I:\LME\TEMPORARIO\Diogo\Automacao\controlista\leituras_ultimo_ponto.txt', 'w', encoding='utf-8')

		if endGpib != '0':
			file = open('I:\LME\TEMPORARIO\Diogo\Automacao\controlista\leituras_ultimo_ponto1.txt', 'w', encoding='utf-8')


		for leitura in range(int(quantasLeituras)):
			
			dmm34970A.write('ROUT:SCAN (@{})'.format(qualCanal))
			a = float(dmm34970A.query('read?'))
			a = str(round(a/float(ajustaCasas),12))
			a = a.replace('.', ',')

			indice = str(leitura+1)
			indice = '[{}/{}]'.format(indice, quantasLeituras)
			print('{}\t{}'.format(indice, a))
			file.write(a+'\n')
			fileCompleta.write(a+'\n')
		
		file.close()
		fileCompleta.write('\n')

		dmm34970A.control_ren(6)
		
		if endGpib == '0':
			keyboard.press_and_release('ctrl + 0')

		if endGpib != '0':
			keyboard.press_and_release('ctrl + 9')
		
		print('\a\a\a\a\n> Próximo ponto? s/n')
		a = input('> ').lower()
		if a == 'n':
			fileCompleta.close()
			break

		print('\a\a\a\a\n> Mesma configuração? s/n')
		mesmoSetup = input('> ').lower()

#########################################################################
# 4338A
#########################################################################

def leitura4338(endGpib, endereco):

	rm = visa.ResourceManager()
	milohm4338 = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')
	# milohm4338.write('*CLS')
			
	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'

	while True:
		
		if mesmoSetup == 'n':

			setup = defineSetup(ajustaCasas=True)
			quantasLeituras = setup['quantasLeituras']
			ajustaCasas = setup['ajustaCasas']
			unidadeSelecionada = setup['unidadeSelecionada']
			
		print('\n[4338A] '+quantasLeituras+'≡ '+unidadeSelecionada)

		# milohm4338.write('*CLS')

		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []


		for leitura in range(int(quantasLeituras)):
			

			# milohm4338.write('*CLS')
			# milohm4338.write('TRIG')
			# milohm4338.write('TRIG:Sour')
			# milohm4338.write('init')
			# milohm4338.write('init:cont')
			# milohm4338.write('abor')
			# milohm4338.query('*TRG')
			a = milohm4338.query('fetc?')
			# print(a)
			# input('.')

			a = a.split(',')[1]
			# a = a.split('E')[0]
			a = float(a)
			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, milohm4338, endGpib)

#########################################################################
# E4980A
#########################################################################

def leituraE4980A():

	print('\n## Leituras com E4980A ##')

	rm = visa.ResourceManager()
	#listaInstrumentos()

	endGpib = '0'
	endGpib = input('Qual interface GPIB? \n> ')

	endereco = input('Qual endereço GPIB?\n> ')
	if endereco[0] == 'n':
		listaInstrumentos(endGpib)
		endereco = input('Qual endereço GPIB?\n> ')
	lcrE4980A = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')
		
	
	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	print(horaData)
	
	if endGpib == '0':
		fileCompleta = open('I:\LME\TEMPORARIO\Diogo\Automacao\controlista\Leituras_{}.txt'.format(horaData), 'w', encoding='utf-8')

	if endGpib != '0':
		fileCompleta = open('I:\LME\TEMPORARIO\Diogo\Automacao\controlista\Leituras1_{}.txt'.format(horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'

	
	while True:
		
		if mesmoSetup == 'n':

			print('\n> Primeira ou Segunda medição?')
			qualMedicao = input('> ').lower()
			
			print('\n> Qual unidade no Autolab?')
			print('> 0 -> pico')
			print('> 1 -> nano')
			print('> 2 -> micro')
			print('> 3 -> mili')
			print('> 4 -> sem multiplicador')
			print('> 5 -> kilo')
			print('> 6 -> Mega')
			print('> 7 -> Giga')
			ajustaCasas = input('> ')
		

			if ajustaCasas == '0' or ajustaCasas == 'pico' or ajustaCasas == 'p':
				unidadeSelecionada = 'pico'
				ajustaCasas = 1/1000_000_000_000
		
			if ajustaCasas == '1' or ajustaCasas == 'nano' or ajustaCasas == 'n':
				unidadeSelecionada = 'nano'
				ajustaCasas = 1/1000_000_000
			
			if ajustaCasas == '2' or ajustaCasas == 'micro' or ajustaCasas == 'u':
				unidadeSelecionada = 'micro'
				ajustaCasas = 1/1000_000
			
			if ajustaCasas == '3' or ajustaCasas == 'mili' or ajustaCasas == 'm':
				unidadeSelecionada = 'mili'
				ajustaCasas = 1/1000
			
			if ajustaCasas == '4' or ajustaCasas == 'sem' or ajustaCasas == '':
				unidadeSelecionada = 'sem mult.'
				ajustaCasas = 1

			if ajustaCasas == '5' or ajustaCasas == 'kilo' or ajustaCasas == 'k':
				unidadeSelecionada = 'kilo'
				ajustaCasas = 1000
			
			if ajustaCasas == '6' or ajustaCasas == 'mega' or ajustaCasas == 'M':
				unidadeSelecionada = 'mega'
				ajustaCasas = 1000_000

			if ajustaCasas == '7' or ajustaCasas == 'giga' or ajustaCasas == 'G':
				unidadeSelecionada = 'giga'
				ajustaCasas = 1000_000_000

		
		# print('\n> Unidade selecionada: '+unidadeSelecionada)				
		# print('> Canal selecionado: {}'.format(qualCanal))
		# print('> Número de leituras: '+quantasLeituras)
		# print('> Range selecionado: {}'.format(unidadeSelecionada))


		if qualMedicao == '1' or qualMedicao[0] == 'p':
			qualMedicaoPrint = 'Medida 1'
		if qualMedicao != '1' or qualMedicao[0] == 's':
			qualMedicaoPrint = 'Medida 2'

		print('\n[E4980A] '+quantasLeituras+'≡ '+unidadeSelecionada+' ('+qualMedicaoPrint+')')

		# lcrE4980A.write('*CLS')
		
				
		if endGpib == '0':
			file = open('I:\LME\TEMPORARIO\Diogo\Automacao\controlista\leituras_ultimo_ponto.txt', 'w', encoding='utf-8')

		if endGpib != '0':
			file = open('I:\LME\TEMPORARIO\Diogo\Automacao\controlista\leituras_ultimo_ponto1.txt', 'w', encoding='utf-8')


		for leitura in range(int(quantasLeituras)):
			
			a = lcrE4980A.write('form:asc:long on')
			a = lcrE4980A.query('fetc?')

			if qualMedicao == '1' or qualMedicao[0] == 'p':
				a = a.split(',')[0]
			if qualMedicao != '1' or qualMedicao[0] == 's':
				a = a.split(',')[1]
			
			a = float(a)
			a = str(round(a/float(ajustaCasas),12))
			a = a.replace('.', ',')

			indice = str(leitura+1)
			indice = '[{}/{}]'.format(indice, quantasLeituras)
			print('{}\t{}'.format(indice, a))
			file.write(a+'\n')
			fileCompleta.write(a+'\n')
		
		file.close()
		fileCompleta.write('\n')

		lcrE4980A.control_ren(6)
		
		if endGpib == '0':
			keyboard.press_and_release('ctrl + 0')

		if endGpib != '0':
			keyboard.press_and_release('ctrl + 9')
		
		print('\a\a\a\a\n> Próximo ponto? s/n')
		a = input('> ').lower()
		if a == 'n':
			fileCompleta.close()
			break

		print('\a\a\a\a\n> Mesma configuração? s/n')
		mesmoSetup = input('> ').lower()

#########################################################################
# E4980A
#########################################################################

def controla33500(endGpib, endereco):

	print('\n## Controla 33500B ##')

	rm = visa.ResourceManager()
	gerador = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')
	gerador.write('OUTP1 OFF')
	gerador.write('OUTP2 OFF')

	print('Aplicar qual forma de onda?')
	formaOnda = input('> ').lower()
	if formaOnda == '' or formaOnda[0] == 's':
		gerador.write('SOUR1:FUNC SIN')
	else:
		gerador.write('SOUR1:FUNC SQU')

	print('Qual impedância? (Ohms)')
	impedancia = input('> ').lower()
	gerador.write('OUTP1:LOAD {}'.format(impedancia))

	print('Vrms, Vpp ou dbm?')
	tipoSinal = input('> ')
	gerador.write('SOUR1:VOLT:UNIT {}'.format(tipoSinal))
	
	while True:

		print('\nQual amplitude do sinal? ('+tipoSinal+')')
		qualAmplitude = input('> ').replace(',','.')
		if qualAmplitude == '':
			amplitude = amplitude
		else:
			amplitude = qualAmplitude
		gerador.write('SOUR1:VOLT {}'.format(amplitude))
		
		print('Qual frequência? (Hz)')
		frequencia = input('> ').replace('Hz', '').replace('HZ', '').replace('hz', '')
		frequencia = frequencia.replace(',','.')

		frequencia = frequencia.replace('m','e-3')
		frequencia = frequencia.replace('k','e3')
		frequencia = frequencia.replace('M','e6')
		frequencia = frequencia.replace('G','e9')

		if frequencia == 'voltar':
			gerador.write('OUTP1 OFF')
			gerador.write('OUTP2 OFF')
			break
		
		if frequencia == '':
			gerador.write('OUTP1 OFF')
			gerador.write('OUTP2 OFF')

		else:
			gerador.write('SOUR1:FREQ {}'.format(frequencia))
			gerador.write('OUTP1 ON')
			gerador.write('OUTP2 OFF')


#########################################################################
# B2987A
#########################################################################

def leituraB2987A(endGpib, endereco):

	rm = visa.ResourceManager()
	B2987A = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')

	input('\n> Configure o multímetro e pressione ENTER')
	
	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	print('Qual grandeza será medida?')
	print('source | i | V | ohm')
	qualGrandeza = input('> ').lower()

	mesmoSetup = 'n'

	while True:
		
		if mesmoSetup == 'n':
		
			setup = defineSetup(ajustaCasas=True)
			ajustaCasas = setup['ajustaCasas']
			quantasLeituras = setup['quantasLeituras']
			unidadeSelecionada = setup['unidadeSelecionada']

		print('[B2987A] '+quantasLeituras+'≡ '+unidadeSelecionada +' '+ qualGrandeza)
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		for leitura in range(int(quantasLeituras)):
			
			if qualGrandeza == 'source':
				a = float(B2987A.query('MEAS:sour?'))

			if qualGrandeza == 'i':
				# a = float(B2987A.query('read:scal:curr?'))
				# a = float(B2987A.query('FETC:curr?'))
				a = float(B2987A.query('MEAS:CURR?'))

			if qualGrandeza == 'ohm':
				a = float(B2987A.query('MEAS:res?'))

			if qualGrandeza == 'v':
				a = float(B2987A.query('MEAS:volt?'))

			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, B2987A, endGpib)


#########################################################################
# NI9225
#########################################################################

def leituraNI9225():

	
	print('Salvar leituras no arquivo 0 ou 1?')
	endGpib = input('> ')

	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'

	while True:
		
		if mesmoSetup == 'n':
		
			print('A placa está em qual slot?')
			qualSlot = input('> ')

			print('Coletar leituras de qual canal?')
			qualCanal = input('> ')

			faixa = 300 ## config do manual da NI
			
			## o manual de calibração da NI pede 50_000 para rotina_rate e 4167 samples
			## utilizamos a configuração abaixo porque ela foi suficiente para pegarmos
			## 10 ondas, considerando que o sinal aplicado será de 60 Hz
			## para outras frequências calcular o rate para ter no mínimo 10 ondas
			rotina_rate = 25_000
			rotina_samples = 5000
			
			setup = defineSetup()
			ajustaCasas = 1
			quantasLeituras = setup['quantasLeituras']


		print('[NI9225] '+quantasLeituras+'≡ ')
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		for leitura in range(int(quantasLeituras)):
			
			with nidaqmx.Task() as task:
  
				# task.ai_channels.add_ai_voltage_rms_chan(
				task.ai_channels.add_ai_voltage_chan(
					'cDAQ1Mod{}/ai{}'.format(qualSlot, qualCanal), #physical_channel
					#name_to_assign_to_channel = '',
					terminal_config = nidaqmx.constants.TerminalConfiguration.DIFFERENTIAL, #config vinda do manual da NI
					# terminal_config = nidaqmx.constants.TerminalConfiguration.DEFAULT,
					min_val= -faixa,
					max_val= faixa,
					units = nidaqmx.constants.VoltageUnits.VOLTS,
					#custom_scale_name = ''
				)

				task.timing.cfg_samp_clk_timing(
					rate = rotina_rate,
					#source = '',
					#active_edge = nidaqmx.constants.Edge.FALLING,
					#active_edge = nidaqmx.constants.Edge.RISING,
					sample_mode = nidaqmx.constants.AcquisitionType.FINITE,
					# sample_mode = nidaqmx.constants.AcquisitionType.CONTINUOUS,
					samps_per_chan = rotina_samples
				)
					
				placa_leitura_lst = task.read(number_of_samples_per_channel = rotina_samples, timeout = 500.0)
				
				## para verificar no excel se a forma de onda está sendo
				## coletada corretamente para a frequência aplicada
				# leiturasPonto_teste = open('./leituras/leituras_teste.txt', 'w', encoding='utf-8')
				# leiturasPonto_teste.write(str(placa_leitura_lst).replace('[','').replace(']','').replace(', ','\n').replace('.',','))
				# leiturasPonto_teste.close()
				# input('.')


				## aqui tratamos o conjunto de leituras para calcular
				## a tensão de pico a pico e a partir daí calcular
				## a tensão RMS
				placa_leitura_lst.sort(reverse = True)
				leitura_max_lst = []
				leitura_min_lst = []
				numero_de_picos = int(rotina_rate/rotina_samples)
				for i in range(1, numero_de_picos):
					leitura_max_lst.append(placa_leitura_lst[i])
					leitura_min_lst.append(placa_leitura_lst[-i])

				leitura_max = statistics.mean(leitura_max_lst)
				leitura_min = abs(statistics.mean(leitura_min_lst))
				leitura_picoApico = leitura_max + leitura_min
				
				a = leitura_picoApico / (2 * np.sqrt(2) )

			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, 'placa NI', endGpib)


#########################################################################
# NI9227
#########################################################################

def leituraNI9227():

	print('Salvar leituras no arquivo 0 ou 1?')
	endGpib = input('> ')
	
	agora = datetime.datetime.now()
	horaData = '{}_{}_{}__{}h{}m'.format(agora.day, agora.month, agora.year, agora.hour, agora.minute)
	leituras_txt = open('./leituras/leituras{}_{}.txt'.format(endGpib, horaData), 'w', encoding='utf-8')

	mesmoSetup = 'n'

	while True:
		
		if mesmoSetup == 'n':
		
			print('A placa está em qual slot?')
			qualSlot = input('> ')

			print('Coletar leituras de qual canal?')
			qualCanal = input('> ')

			faixa = 5 ## config do manual da NI
			## o manual de calibração da NI pede 50_000 para rotina_rate e 5_000 samples
			## utilizamos a configuração abaixo porque ela foi suficiente para pegarmos
			## 10 ondas, considerando que o sinal aplicado será de 60 Hz
			## para outras frequências calcular o rate para ter no mínimo 10 ondas
			rotina_rate = 25_000
			rotina_samples = 5000
			
			setup = defineSetup()
			ajustaCasas = 1
			quantasLeituras = setup['quantasLeituras']


		print('[NI9227] '+quantasLeituras+'≡ canal '+qualCanal)
		leiturasPonto_txt = open('./leituras/leituras_ultimo_ponto{}.txt'.format(endGpib), 'w', encoding='utf-8')
		setPonto = []

		for leitura in range(int(quantasLeituras)):
			
			with nidaqmx.Task() as task:
  
				task.ai_channels.add_ai_current_chan(
					'cDAQ1Mod{}/ai{}'.format(qualSlot, qualCanal), #physical_channel
					#name_to_assign_to_channel = '',
					#terminal_config=<TerminalConfiguration.DEFAULT: -1>
					terminal_config = nidaqmx.constants.TerminalConfiguration.DIFFERENTIAL, #config vinda do manual da NI
					# terminal_config = nidaqmx.constants.TerminalConfiguration.DEFAULT,
					min_val= -faixa,
					max_val= faixa,
					units = nidaqmx.constants.CurrentUnits.AMPS,
					shunt_resistor_loc = nidaqmx.constants.CurrentShuntResistorLocation.INTERNAL,
					# ext_shunt_resistor_val=,
					#custom_scale_name = ''
				)
				
				
				task.timing.cfg_samp_clk_timing(
					rate = rotina_rate,
					#source = '',
					#active_edge = nidaqmx.constants.Edge.FALLING,
					#active_edge = nidaqmx.constants.Edge.RISING,
					sample_mode = nidaqmx.constants.AcquisitionType.FINITE,
					# sample_mode = nidaqmx.constants.AcquisitionType.CONTINUOUS,
					samps_per_chan = rotina_samples
				)
					
				placa_leitura_lst = task.read(number_of_samples_per_channel = rotina_samples, timeout = 500.0)
				
				## para verificar no excel se a forma de onda está sendo
				## coletada corretamente para a frequência aplicada
				# leiturasPonto_teste = open('./leituras/leituras_teste.txt', 'w', encoding='utf-8')
				# leiturasPonto_teste.write(str(placa_leitura_lst).replace('[','').replace(']','').replace(', ','\n').replace('.',','))
				# leiturasPonto_teste.close()
				# input('.')

				## aqui tratamos o conjunto de leituras para calcular
				## a tensão de pico a pico e a partir daí calcular
				## a tensão RMS
				placa_leitura_lst.sort(reverse = True)
				leitura_max_lst = []
				leitura_min_lst = []
				numero_de_picos = int(rotina_rate/rotina_samples)
				for i in range(1, numero_de_picos):
					leitura_max_lst.append(placa_leitura_lst[i])
					leitura_min_lst.append(placa_leitura_lst[-i])

				leitura_max = statistics.mean(leitura_max_lst)
				leitura_min = abs(statistics.mean(leitura_min_lst))
				leitura_picoApico = leitura_max + leitura_min
				
				a = leitura_picoApico / (2 * np.sqrt(2) )

			processaLeitura(a, ajustaCasas, leitura, quantasLeituras, leiturasPonto_txt, leituras_txt, setPonto)
		
		mesmoSetup = proximoPonto(leiturasPonto_txt, leituras_txt, 'placa NI', endGpib)



#########################################################################
# Livre - apenas conecta e manda os comandos como digitados
#########################################################################

def livre():

	rm = visa.ResourceManager()
	#listaInstrumentos()

	endGpib = '0'
	endGpib = input('Qual interface GPIB? \n> ')

	endereco = input('Qual endereço GPIB?\n> ')
	if endereco[0] == 'n':
		listaInstrumentos(endGpib)
		endereco = input('Qual endereço GPIB?\n> ')
	instrumento = rm.open_resource('GPIB{}::{}::INSTR'.format(endGpib, endereco), timeout=200000, read_termination='\n')
	instrumento.control_ren(6)
		
	print('\n## Conectado ##')

	while True:

		print('\nDigite o comando completo a ser enviado')
		a = input('> ')

		if a.lower() == 'sair':
			break

		tipo = a.split(',')[0]

		print(tipo)
		print(a)

		if tipo == 'w':
			instrumento.write(a.split(',')[1])
			r = instrumento.write(a.split(',')[1])
		
		if tipo == 'q':
			r = instrumento.query(a.split(',')[1])

		if tipo == 'r':
			r = instrumento.read(a.split(',')[1])

		instrumento.control_ren(6)

		print (r)


#########################################################################
# main
#########################################################################
def main(): 
	while True:

		
		# print(Style.RESET_ALL)
		ctypes.windll.kernel32.SetConsoleTitleW("controlista")
		print('\n#########################################')
		print('0. Instrumentos conectados')
		print('1. 57xxx')
		print('2. 55xx')
		print('3. 3458A')
		print('4. 5335A')
		print('5. PM6304')
		print('6. 8508A')
		print('7. 34410A')
		print('8. 8846A / 8845A / 88XXX')
		print('9. 5790A')
		print('10. 34420A')
		print('11. QuadTech 1920')
		print('12. 34970A')
		print('13. 7561')
		print('14. 4339B / 4338A / 4263B')
		print('15. E4980A')
		print('16. 33500B')
		print('17. B2987A')
		print('18. NI9225')
		print('19. NI9227')
		qual1 = input('> ').lower().replace(' ', '')

		ctypes.windll.kernel32.SetConsoleTitleW(qual1.upper())

		if qual1 == '0' or qual1 == 'lista':
			listaInstrumentos()
			main()
		
		if qual1 != '0' and qual1 != '18' and qual1 != 'ni9225' and qual1 != '19' and qual1 != 'ni9227':
			endGpib, endereco = pegaEnderecos()

		if qual1 == '1' or qual1[0:4] == '5700' or qual1[0:4] == '5720':
			cal5700AsII(endGpib, endereco)

		if qual1 == '2' or qual1[0:4] == '5520' or qual1[0:4] == '5500' or qual1[0:4] == '5522' or qual1[0:4] == '5800':
			cal55XXX(endGpib, endereco)
		
		if qual1 == '3' or qual1 == '3458' or qual1 == '3458a':
			leitura3458A(endGpib, endereco)
			
		if qual1 == '4' or qual1 == '5335':
			input('REINICIE o contador e pressione ENTER')
			leitura3553A(endGpib, endereco)

		if qual1 == '5' or qual1 == '6304' or qual1 == 'pm6304':
			livrePM6304(endGpib, endereco)
		
		if qual1 == '6' or qual1 == '8508' or qual1 == '8508a':

			input('Faça o zero do instrumento e pressione ENTER')		
			input('Faça a configuração adequada no painel \ndo instrumento e pressione ENTER')
			leitura8508A(endGpib, endereco)

		if qual1 == '7' or qual1[0:5] == '34401' or qual1[0:5] == '34411':
			leitura344XXX(endGpib, endereco)			
	
		if qual1 == '8' or qual1[0:4] == '8846' or qual1[0:4] == '8845':
			leitura88XXX(endGpib, endereco)

		if qual1 == '9' or qual1[0:4] == '5790':
			input('Faça a configuração adequada e pressione ENTER')
			leitura5790A(endGpib, endereco)

		if qual1 == '10' or qual1 == '34420':
			input('Faça a configuração adequada e pressione ENTER')
			leitura34420A(endGpib, endereco)

		if qual1 == '11' or qual1 == 'quadtech' or qual1 == 'quad' or qual1 == '1920':
			leituraQuadtech(endGpib, endereco)

		if qual1 == '12' or qual1 == '34970':
			leitura34970A(endGpib, endereco)

		if qual1 == '13':
			print('Selecione a grandeza a ser medida, ')
			print('faça a configuração adequada, deixe o trigger')
			input('no modo AUTOMÁTICO/INTERNO e pressione ENTER')
			leitura7561(endGpib, endereco)

		if qual1 == '14' or qual1 == '4338' or qual1[0:4] == '4263' or qual1[0:4] == '4339':
			print('Faça a configuração adequada, deixe o trigger')
			input('no modo AUTOMÁTICO/INTERNO e pressione ENTER')
			leitura4338(endGpib, endereco)

		if qual1 == '15' or qual1 == '4980':
			print('Faça a configuração adequada, deixe o trigger')
			input('no modo AUTOMÁTICO/INTERNO e pressione ENTER')
			leituraE4980A(endGpib, endereco)

		if qual1 == '16' or qual1 == '33500' or qual1 == '33500b':
			controla33500(endGpib, endereco)

		if qual1 == '17' or qual1 == 'b2987a' or qual1 == 'b2987':
			leituraB2987A(endGpib, endereco)

		if qual1 == '18' or qual1 == 'ni9225':
			leituraNI9225()
		
		if qual1 == '19' or qual1 == 'ni9227':
			leituraNI9227()

		if qual1 == 'livre':
			livre(endGpib, endereco)
		
		if qual1 == 'sair':
			break

###############################################################
main()

