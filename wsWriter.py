#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@TIGR Metrics@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@||
'''  #																	||
---  #																	||
<(META)>:  #															||
	DOCid:   #						||
	name: TIGR   #					||
	description: >  #													||
		Calculate Quality Load Factor from Recent Data.  Write a tab #	||
		for each Optics Production Type Code pulling PTC overrides
		from excel config file and add any additional PTC codes
	expirary: {expiration}  #											||
	version: {version}  #												||
	path: {LEXIvrs}  #												||
	outline: {outline}  #												||
	authority: document|this  #											||
	security: sec|lvl2  #												||
	<(WT)>: -32  #														||
''' #																	||
# -*- coding: utf-8 -*-#												||
#===============================Core Modules====================================||
from os.path import abspath, dirname, join
from copy import deepcopy
import logging
from logging import getLogger, basicConfig, DEBUG
from logging.config import fileConfig
#===============================3rd Party Modules===============================||
import json
#================================Pheonix Modules================================||
from get import getYAML
import xl, util
#===============================================================================||
#fileConfig('logging_config.ini')
#basicConfig(level=DEBUG, filename='log.log')
basicConfig(filename='example.log',level=DEBUG)
logger = getLogger()
here = join(dirname(__file__),'')#						||
there = abspath(join('../../..'))#						||set path at pheonix level
version = '0.0.0.0.0.0'#												||
#===============================================================================||
def mapDistWS(plants, ptcs, qlfs, cfgs=None):
	''
#	print('Map Distribution WS')
	if cfgs == None:
		cfg = getYAML('{0}z-data_/qlf.yaml'.format(here))
	else:
		cfg = getYAML('{0}'.format(cfgs['qlf']))
	if cfgs == None:
		xlcfg = getYAML('{0}z-data_/xl.yaml'.format(here))
	else:
		xlcfg = getYAML('{0}'.format(cfgs['xl']))
	sumifFX = xlcfg['formulas']['sumif']['fx']#							||
	sumifsFX = xlcfg['formulas']['sumifs']['fx']#							||
	tmplt = cfg['TMPLTs']['sheets']['Dist']
	body, r = [], 2
	sr, er = 2, 10000
	for ptc in sorted(ptcs):#													||
		if '/' in ptc:
			ptc = ptc.replace('/','')
		c = 0
		plants.sort()
		for plant in plants:
			datar = []#													||
			if ptc == 'ttl' or ptc == '':
				continue
#			if plant not in qlfs[ptc]['QLF'].keys():
#				datar = [ptc, plant, 0, 0, 0]
#			else:
			sumifDT = {'cr0': "'ProductMix'!A{0}".format(sr),
				 			'cr1': "A{0}".format(er),
							'cr2': "'Distribution'!A{0}".format(r),
							'cr3': "'ProductMix'!C{0}".format(sr),
							'cr4': "C{0}".format(er)}
			sumif = sumifFX.format(**sumifDT)
			sumifsDT = {'cr0': "'ProductMix'!A{0}".format(sr),
							'cr1': "A{0}".format(er),
							'cr2': "'Distribution'!A{0}".format(r),
							'cr3': "ProductMix!C{0}".format(sr),
							'cr4': "C{0}".format(er),
							'cr5': "'ProductMix'!B{0}".format(sr),
							'cr6': "B{0}".format(er),
							'cr7': "'Distribution'!B{0}".format(r)}
			sumifs = sumifsFX.format(**sumifsDT)
			sumif = sumif.replace('=', '')
			fx = '=IF(ISERROR(C{0}/{1}),0,C{2}/{3})'.format(r,sumif,r,sumif)
			datar = [ptc, plant, sumifs, fx]
			body.append(datar)#											||
			r += 1#														||
			c += 1#														||
#	print('Dist Body',body)
	fdmap, stmap, szmap = xl.creWSTable(body, tmplt)
#	print('StyleMap', stmap)
	return fdmap, stmap, szmap#											||
def mapPTCWS(ptc, qlf, parts, prtes, plants, vols, cfgs=None):
#	print('Map PTC',ptc)
#	log.WORKLOG('MapPTC',ptc,qlf)
	if cfgs == None:
		cfg = getYAML('{0}z-data_/qlf.yaml'.format(here))
	else:
		cfg = getYAML('{0}'.format(cfgs['qlf']))
	if cfgs == None:
		xlcfg = getYAML('{0}z-data_/xl.yaml'.format(here))
	else:
		xlcfg = getYAML('{0}'.format(cfgs['xl']))
	sizes = cfg['TMPLTs']['sheets']['PTC']['size']
	styles = cfg['TMPLTs']['sectns']['PTC']['styles']
	stmap = {'background': styles['background']}#						||
	sectnTMPLT = cfg['TMPLTs']['sectns']['PTC']['map']#					||
	hdr, ftr, fdmap = sectnTMPLT['header'], sectnTMPLT['footer'], {}#	||
	fx0 = xlcfg['formulas']['sumifs']['fx']#							||
#	print('HEADer',hdr)
	bhead = len(hdr.keys())+len(prtes['RTE1'].keys())#					||
	bhead += len(ftr.keys())+2#											||
	coln, acnt = 1, 0#													||
	srown = rown = len(plants)*bhead-4
	name_col, plant_col = xl.getLetter(coln), xl.getLetter(coln+1)#		||
	rate_col, rqpcs_col = xl.getLetter(coln+2), xl.getLetter(coln+3)#	||
	rhrs_col, gpcs_col = xl.getLetter(coln+4), xl.getLetter(coln+5)#	||
	tpcs_col, spcs_col = xl.getLetter(coln+6), xl.getLetter(coln+7)#	||
	srate_col, qlf_col = xl.getLetter(coln+8), xl.getLetter(coln+9)#	||
	qhrs_col, chrs_col = xl.getLetter(coln+10), xl.getLetter(coln+11)#	||
	for rc in sorted(qlf.keys()):#												||routing code
		if rc == 'qtys':
			continue
#		print('Route Code',rc)
#		print('qlfrc',qlf[rc])
		for pc in sorted(qlf[rc].keys()):#										||plant code
			asrown = rown
			if pc == 'dist':
				continue
			for sc in sorted(qlf[rc][pc].keys()):#								||sequence code
				body, q = [], qlf[rc][pc][sc]#							||
				bsrown = rown + len(hdr.keys())#						||
				r = rown
#				print('PTCWS RTE', q['RTE'])
				berown = bsrown + len(q['RTE'].keys())-1#					||
#				rteseq = reversed(list(q['RTE'].keys()))
				logging.debug(q['RTE'])
#				for step in sorted(q['RTE'].keys(), reverse=True):#rteseq.sort():#									||
				for step in q['RTE'].keys():
					if 'Name' not in q['RTE'][step].keys():#			||
						continue#										||
					qs = q['RTE'][step]#								||
					try:
						rate = qs['Rate']
					except:#											||
						rate = '---'
					try:
						rhrs = qs['RunHours']#							||
					except:
						rhrs = 0
					if qs['Name'][:3] == 'CP ':
						cp = rhrs
					else:
						cp = 0
					ct = float(rhrs)*3600/float(q['PcsReq'])
#					print('CT',ct, 'QLF', qs['QLFRate'])
					try:
						ect = float(ct)*float(qs['QLFRate'])
					except:
						ect = 1.0
					body.append([qs['Name'], pc, rate, q['PcsReq'],rhrs,
									qs['GoodQty'], qs['TotalQty'],
									qs['ScrapPcs'], qs['ScrapRate'],
									qs['QLFRate'], qs['QLFHrs'], cp, ct, ect])
					r += 1
				t = {'ptc': ptc, 'cr00': rqpcs_col+str(bsrown),
						'num_parts': '0', 'num_orders': '0','rte_code': rc}#	||
				head = util.dictSub(deepcopy(hdr), t)
				t = {'cr00': rhrs_col+str(bsrown), 'cr01': rhrs_col+str(berown),#	||
					'cr02': gpcs_col+str(bsrown), 'cr03': gpcs_col+str(berown),#	||
					'cr04': tpcs_col+str(bsrown), 'cr05': tpcs_col+str(berown),#	||
					'cr06': qhrs_col+str(bsrown), 'cr07': qhrs_col+str(berown),#	||
					'cr08': chrs_col+str(bsrown), 'cr09': chrs_col+str(berown),#	||
					'cr10': qlf_col+str(berown+2), 'cr11': qlf_col+str(berown+3),#	||
					'cr12': qlf_col+str(berown+3), 'cr13': qlf_col+str(berown+2)}#	||
				foot = util.dictSub(deepcopy(ftr), t)
				params = ['PTC', head, body, foot, acnt, coln, asrown]#	||
				mapp, ecoln, rown, acnt = xl.mapTableAREA(*params)#		||

				fdmap.update(mapp['value'])
				stmap.update(mapp['stmap'])
				acnt += 1#												||
				rown += 1
				asrown = rown
	dsrown, derown = srown, rown
	terms = {'cr0': name_col+str(dsrown),#							||
									'cr1': name_col+str(derown),#	||
													'cr2': '{c}',#	||
								'cr3': '{a}', 'cr4': '{b}',#	||
									'cr5': plant_col+str(dsrown),#	||
									'cr6': plant_col+str(derown),#	||
													'cr7': '{d}',}#	||
	fx0 = fx0.format(**terms)
	rown = 1
	plants.sort()
	for plant in plants:
		srown = rown
		body, rown = [], rown+len(hdr.keys())#							||
		msrown = rown
		for seq in sorted(prtes['RTE1'].keys()):
			t = {'a': '{a}','b': '{b}', 'c': name_col+str(rown),'d': plant_col+str(rown),}#	||
			fxr = fx0.format(**t)
			datar = [prtes['RTE1'][seq], plant, '']
			if plant == 'STARs':
				datar.append(vols[plant]['pcs'])
			else:
				t = {'a': rqpcs_col+str(dsrown),'b': rqpcs_col+str(derown)}#	||
				datar.append(fxr.format(**t))#								||
			t = {'a': rhrs_col+str(dsrown),'b': rhrs_col+str(derown)}#	||
			datar.append(fxr.format(**t))#								||RunHours
			t = {'a': gpcs_col+str(dsrown),'b': gpcs_col+str(derown),}#	||
			datar.append(fxr.format(**t))#								||GoodPcs
			t = {'a': tpcs_col+str(dsrown),'b': tpcs_col+str(derown),}#	||
			datar.append(fxr.format(**t))#								||
			t = {'a': spcs_col+str(dsrown),'b': spcs_col+str(derown),}#	||
			datar.append(fxr.format(**t))#								||
			t = {'a': gpcs_col+str(rown), 'b': spcs_col+str(rown),#		||
											'c': gpcs_col+str(rown)}#	||
			datar.append('=IF({a}=0,0,{b}/{c})'.format(**t))#			||
			t = {'cr0': qlf_col+str(rown+1), 'cr1': srate_col+str(rown)}#	||
			fx = xlcfg['formulas']['add']['fx'].replace('{','{')#		||
			fx = fx.replace('}','}')#									||
			datar.append(fx.format(**t))#								||QLF
			t = {'a': qlf_col+str(rown), 'b': rhrs_col+str(rown)}#		||
			datar.append('=({a}-1)*{b}'.format(**t))#					||QLFHrs
			t = {'a': chrs_col+str(dsrown), 'b': chrs_col+str(derown)}#	||CPHrs
			datar.append(fxr.format(**t))#								||

			fmla = '=IF(B{r0}="CP BUILD", SUM(L{r1}:L{r2})/G{r1}*10000-SUM(P{r1}:P{r2}),'
			fmla += 'IF(B{r0}="CP TEST", SUM(L{r1}:L{r2})/G{r1}*10000-SUM(P{r1}:P{r2}), ""))'
			sugcphrs = fmla.format(**{'r0':rown, 'r1': msrown, 'r2': rown-1})
			print('SUGGEST CPHrs', sugcphrs)
#			print('Rown',rown)
			ctfx = '=IF(E{0}=0,0,F{1}*3600/E{2})'.format(rown, rown, rown)
			ectfx = '=N{0}*K{1}'.format(rown, rown)
#			print('CT', ctfx, 'ECT', ectfx)
			datar += [ctfx, ectfx, sugcphrs]
			body.append(datar)#											||
			rown += 1
		t = {'ptc': ptc, 'cr00': rqpcs_col+str(msrown),
				'cr01': rqpcs_col+str(rown-1), 'num_parts': '0',#	||
					'num_orders': '0', 'rte_code': 'prime'}#	||
#		print('MSROWN',hdr)
		head = util.dictSub(deepcopy(hdr), t)
#		print('Head', hdr)
		t = {'cr00': rhrs_col+str(msrown), 'cr01': rhrs_col+str(rown-1),#	||
			'cr02': gpcs_col+str(msrown), 'cr03': gpcs_col+str(rown-1),#	||
			'cr04': tpcs_col+str(msrown), 'cr05': tpcs_col+str(rown-1),#	||
			'cr06': qhrs_col+str(msrown), 'cr07': qhrs_col+str(rown-1),#	||
			'cr08': chrs_col+str(msrown), 'cr09': chrs_col+str(rown-1),#	||
			'cr10': qlf_col+str(rown+1), 'cr11': qlf_col+str(rown+2),#	||
			'cr12': qlf_col+str(rown+1), 'cr13': qlf_col+str(rown+2)}#	||
		foot = util.dictSub(deepcopy(ftr), t)
		params = ['PTC', head, body, foot, acnt, 1, srown]
		mapp, ecolb, rown, acnt = xl.mapTableAREA(*params)#	||
		rown += 1
		fdmap.update(mapp['value'])#							||
		stmap.update(mapp['stmap'])#							||
		acnt += 1#														||
	stmap['background']['range']['minrow'] = 0#1#						||
	stmap['background']['range']['mincol'] = 0#						||
	stmap['background']['range']['maxrow'] = 0#derown+5
	stmap['background']['range']['maxcol'] = 0#ecolb+8
	sizemaps = {'columnmap': sizes['cols'], 'rowmap': {}}
	for i in range(rown):
		if i not in sizemaps['rowmap'].keys():#							||
			sizemaps['rowmap'][i] = 15.75#								||
#	log.WORKLOG('stmap',stmap,'')
	return fdmap, stmap, sizemaps
def mapProductMixWS(plants, ptcs, n, cfgs=None):
	''
#	print('Map Product Mix WS')
	if cfgs == None:
		cfg = getYAML('{0}z-data_/qlf.yaml'.format(here))
	else:
		cfg = getYAML('{0}'.format(cfgs['qlf']))
	if cfgs == None:
		xlcfg = getYAML('{0}z-data_/xl.yaml'.format(here))
	else:
		xlcfg = getYAML('{0}'.format(cfgs['xl']))
	tmplt = cfg['TMPLTs']['sheets']['ProductMix']
	body, r = [], 2
	n += 1
	ptcs.sort()
	for ptc in ptcs:#													||
		if '/' in ptc:
			ptc = ptc.replace('/', '')
		c = 0
		for plant in plants:
			datar = []#													||
			if ptc == 'ttl' or ptc == '':
				continue
			datar = [ptc, plant, "='{0}'!E{1}".format(ptc, 1+c*n),
						"='{0}'!F{1}".format(ptc, 18+c*n), "='{0}'!G{1}".format(ptc, 4+c*n)]#	||
			fx = '=IF(ISERROR({cr0}/{cr1}*{cr2}),0,{cr0}/{cr1}*{cr2})'#							||Good Hrs
			terms = {'cr0': 'D{0}'.format(r), 'cr1': 'C{0}'.format(r),
						'cr2': 'E{0}'.format(r)}#						||
			fx = fx.format(**terms)
			datar += [fx, "='"+ptc+"'!H"+str(4+c*n)]#						||Total Pcs
			fx = '=IF(ISERROR({cr0}/{cr1}*{cr2}),0,{cr3}/{cr4}*{cr5})'#									||Good Hrs
			terms = {'cr0': 'D{0}'.format(r), 'cr1': 'C{0}'.format(r),
						'cr2': 'G{0}'.format(r), 'cr3': 'D{0}'.format(r),
						'cr4': 'C{0}'.format(r), 'cr5': 'G{0}'.format(r)}#	||
			fx = fx.format(**terms)
			datar += [fx, "='{0}'!K{1}".format(ptc, 4+c*n), "=F{0}/3".format(r),#	||
						"=G{0}/3".format(r), "=F{0}/90".format(r), "=G{0}/90".format(r),
						'=A{0}&B{1}'.format(r,r)]#	||
			body.append(datar)#											||
			r += 1#														||
			c += 1#														||
	fdmap, stmap, szmap = xl.creWSTable(body, tmplt)
	return fdmap, stmap, szmap#											||
def mapQLFRTEsWS(ptcs, processes, plants, cfgs=None):
	''
#	print('MAP QLFRTE Worksheet')
	if cfgs == None:
		cfg = getYAML('{0}z-data_/qlf.yaml'.format(here))
	else:
		cfg = getYAML('{0}'.format(cfgs['qlf']))
	if cfgs == None:
		xlcfg = getYAML('{0}z-data_/xl.yaml'.format(here))
	else:
		xlcfg = getYAML('{0}'.format(cfgs['xl']))

	tmplt = cfg['TMPLTs']['sheets']['QLFRTEs']
	body, sr, n = [], 4, 9+len(processes.keys())
	r0 = 2
	for ptc in sorted(ptcs):#															||
		if '/' in ptc:
			ptc = ptc.replace('/', '')
		if ptc == 'ttl' or ptc == '':
			continue
		c = 0
		for plant in plants:
			r = sr+c*n
			for seq, process in processes.items():
				body.append([ptc, plant, "='{0}'!B{1}".format(ptc, r),
												"='{0}'!K{1}".format(ptc, r),#	||
												"='{0}'!N{1}".format(ptc, r),#	||
												"='{0}'!O{1}".format(ptc, r),#	||
										'=A{0}&B{1}&C{2}'.format(r0,r0,r0)])#	||
				r += 1
				r0 += 1
			c += 1
	fdmap, stmap, szmap = xl.creWSTable(body, tmplt)
	return fdmap, stmap, szmap
def mapCPUpdateWS(qlfs, rtes, cfgs=None):#														||
	''
	#need a list of parts for each route number and PTC
	#create the tab with each part and route and cp step?
	#find the template for routing update
	if cfgs == None:#															||
		cfg = getYAML('{0}z-data_/qlf.yaml'.format(here))#						||
	else:
		cfg = getYAML('{0}'.format(cfgs['qlf']))#								||
	if cfgs == None:
		xlcfg = getYAML('{0}z-data_/xl.yaml'.format(here))#						||
	else:
		xlcfg = getYAML('{0}'.format(cfgs['xl']))#								||
	tmplt = cfg['TMPLTs']['sheets']['CPUpdates']#								||
	body, r = [], 2#															||
	for ptc in rtes.keys():#													||
		if ptc not in rtes.keys():
			continue
		fmla = '''=VLOOKUP(C{r}, '{ptc}'!B{r0}:P{r1}, 15, FALSE)'''
		opts = {'step': 'CP BUILD', 'ptc': ptc, 'r': '{r}', 'r0': 4, 'r1': 17}
		nafmla0 = fmla.format(**opts)
		opts['step'] = 'CP TEST'
		nafmla1 = fmla.format(**opts)
		opts = {'step': 'CP BUILD', 'ptc': ptc, 'r': '{r}', 'r0': 27, 'r1': 40}
		smefmla0 = fmla.format(**opts)
		opts['step'] = 'CP TEST'
		smefmla1 = fmla.format(**opts)
		for rte in rtes[ptc].keys():
			for part in rtes[ptc][rte].keys():#								||
				cpbuildlock, cptestlock = 0, 0
				for step in rtes[ptc][rte][part]:
					if 'CP BUILD' == step[5]:
						body.append([part, rte, step[5], step[6],
											nafmla0.format(**{'r': r}),
											smefmla0.format(**{'r': r}), ''])#	||
						cpbuildlock = 1
						r += 1
					if 'CP TEST' == step[5]:
						body.append([part, rte, step[5], step[6],
											nafmla1.format(**{'r': r}),
											smefmla1.format(**{'r': r}), ''])#	||
						cptestlock = 1
						r += 1
				if cpbuildlock == 0:
					for step in rtes[ptc][rte][part]:
						if 'HEAT SINK CONTAINMENT TEST' == step:
							body.append([part, rte, 'CP BUILD', 'None',
											nafmla0.format(**{'r': r}),
											smefmla0.format(**{'r': r}), ''])#	||
							r += 1
				if cptestlock == 0:
					for step in rtes[ptc][rte][part]:
						if 'CLEAN / INSPECT' == step:
							body.append([part, rte, 'CP TEST', 'None',
											nafmla1.format(**{'r': r}),
											smefmla1.format(**{'r': r}), ''])#	||
							r += 1
#	logging.debug(['CP Body',body])
	fdmap, stmap, szmap = xl.creWSTable(body, tmplt)
	return fdmap, stmap, szmap
def writeWS(ptc, fdmap, wb, stylemap=None, sizemaps=None):
	''
	sheet = wb.creWS(ptc)#														||
	if stylemap != None:
		xl.styleSheet(sheet, stylemap)#											||
	if sizemaps != None:
		if 'columnmap' in sizemaps.keys():
			wb.setColumnWidths(sheet, sizemaps['columnmap'])
		if 'rowmap' in sizemaps.keys():
			wb.setRowHeights(sheet, sizemaps['rowmap'])
#	print('writeWS',ptc,'FDmap', fdmap)
	wb.mapDict2Cells(sheet, fdmap)
