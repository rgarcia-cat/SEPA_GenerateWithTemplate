'use strict'

const EventEmitter = require('events').EventEmitter;
var XLSX = require('xlsx');
var mysql = require('mysql');
var utf8 = require('utf8');
var xml = require('xml');

var config = require("./config/config");

String.prototype.NameCase = function() {
   	var splitStr = this.toLowerCase().split(' ');
   	for (var i = 0; i < splitStr.length; i++) {
 		if (splitStr[i].length > 1)
    	splitStr[i] = splitStr[i].charAt(0).toUpperCase() + splitStr[i].substring(1);
   	}
   	return splitStr.join(' ');
}

function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

class utfFix extends EventEmitter{
	constructor(){
		super();
		this.arrChange = [
						{ database: 'Ã‰', real:'É' },
						{ database: 'Ãˆ', real:'È' },
						{ database: 'Ã“', real:'Ó' },
						{ database: 'Ã’', real:'Ò' }
						]
	}
	fixed(str){
		for (let i=0; i < this.arrChange.length; i++)
			str = str.replace(this.arrChange[i].database,utf8.encode(this.arrChange[i].real));

		return str;
	}
}

class RegManager extends EventEmitter{
  	constructor(){
	    super();
	    this.status = "start";
	    this.bFirstFix = false;
  	}
	tractarRegistre(r){
		if(!Sepa.bCheck){
			var money = parseFloat(r.Import.replace('€', '').replace(',','.'));
			Sepa.iAmount += money;
			Sepa.iCounter++;
			var aData = r.DataCreacio.split(' ');
			Sepa.sDataCreacio = aData[0].substring(6,10) + '-' + aData[0].substring(3,5) + '-' + aData[0].substring(0,2) + 'T' + aData[1] + ':00';
			aData = r.DataCarrec.split('/');
			Sepa.sDataRemesa = aData[2] + '-' + aData[1] + '-' + aData[0];
			Sepa.sIdRemesa = String(r.IdRemesa).replace(/-/g, '');
			this.emit('wrote');
		}else{
			this.addPayment(r);
		}
	}
	
	charFix(strng){
		var r=strng.toLowerCase();
            r = r.replace(new RegExp(/[àáâãäå]/g),"a");
            r = r.replace(new RegExp(/[èéêë]/g),"e");
            r = r.replace(new RegExp(/[ìíîï]/g),"i");
            r = r.replace(new RegExp(/ñ/g),"n");     
            r = r.replace(new RegExp(/ç/g),"c");
            r = r.replace(new RegExp(/'/g),"c");            
            r = r.replace(new RegExp(/[òóôõö]/g),"o");
            r = r.replace(new RegExp(/[ùúûü]/g),"u");  
            r = r.replace(new RegExp(/[&]/g),""); 
 		return r;
	}
	addPayment(r){
		if(this.bFirstFix){
			if(!Sepa.bEndPayment){
				var strngId = String(r.IdPagament).replace(/-/g, '');
				var strngNom = r.NomPer + ' ' + r.CognomPer + ' ' + r.NomOrg;
				var aNom = strngNom.split(' ');
				strngNom = '';
				for(var i=0; i<aNom.length; i++){
					if(aNom[i] != null && aNom[i].length != 0){
						aNom[i] = this.charFix(aNom[i]).trim();
						aNom[i] = aNom[i].substring(0,1).toUpperCase() + aNom[i].substring(1, aNom[i].length).toLowerCase();
						strngNom += aNom[i] + ' ';
					}
				}
				var aData = r.Data.split('/');

				console.log('\t\t\t<DrctDbtTxInf>');
				console.log('\t\t\t\t<PmtId>');
				console.log('\t\t\t\t\t<InstrId>' + strngId.trim() + '</InstrId>');
				console.log('\t\t\t\t\t<EndToEndId>' + strngId.trim() + '</EndToEndId>');
				console.log('\t\t\t\t</PmtId>');
				console.log('\t\t\t\t<InstdAmt Ccy="' + config.Money.trim() + '">' + parseFloat(r.Import.replace('€', '').replace(',','.')).toFixed(2) + '</InstdAmt>');
				console.log('\t\t\t\t<DrctDbtTx>');
				console.log('\t\t\t\t\t<MndtRltdInf>');
				console.log('\t\t\t\t\t\t<MndtId>' + r.Mandat.trim() + '</MndtId>');
				console.log('\t\t\t\t\t\t<DtOfSgntr>' + aData[2].trim() + '-' +  aData[1].trim() + '-' + aData[0].trim() + '</DtOfSgntr>');
				console.log('\t\t\t\t\t\t<AmdmntInd>' + 'false' + '</AmdmntInd>');
				console.log('\t\t\t\t\t</MndtRltdInf>');
				console.log('\t\t\t\t</DrctDbtTx>');
				console.log('\t\t\t\t<DbtrAgt>');
				console.log('\t\t\t\t\t<FinInstnId>');
				console.log('\t\t\t\t\t\t<Othr>');
				console.log('\t\t\t\t\t\t\t<Id>' + config.PersonId.trim() + '</Id>');
				console.log('\t\t\t\t\t\t</Othr>');
				console.log('\t\t\t\t\t</FinInstnId>');
				console.log('\t\t\t\t</DbtrAgt>');
				console.log('\t\t\t\t<Dbtr>');
				console.log('\t\t\t\t\t<Nm>' + strngNom.trim() + '</Nm>');
				console.log('\t\t\t\t</Dbtr>');
				console.log('\t\t\t\t<DbtrAcct>');
				console.log('\t\t\t\t\t<Id>');
				console.log('\t\t\t\t\t\t<IBAN>' + r.Iban.trim() + r.CCC.trim() + '</IBAN>');
				console.log('\t\t\t\t\t</Id>');
				console.log('\t\t\t\t</DbtrAcct>');
				console.log('\t\t\t\t<RmtInf>');
				console.log('\t\t\t\t\t\t<Ustrd>' + config.Concept.trim() + '</Ustrd>');
				console.log('\t\t\t\t</RmtInf>');
				console.log('\t\t\t</DrctDbtTxInf>');
			}
		}
		this.bFirstFix = true;
		this.emit('wrote');
	}
}


class excelWeb extends EventEmitter {
  constructor(file){
    super();
    this.indicadors = [ 'A' ];
    this.file = file;
    this.status = "start";
	this.columnes = [
					{ column: 'A', valor: 'IdPagament'},
					{ column: 'B', valor: 'Import'},
					{ column: 'C', valor: 'Mandat'},
					{ column: 'D', valor: 'Data'},
					{ column: 'E', valor: 'Iban'},
					{ column: 'F', valor: 'CCC'},
					{ column: 'G', valor: 'IdRemesa'},
					{ column: 'H', valor: 'DataCreacio'},
					{ column: 'I', valor: 'DataCarrec'},
					{ column: 'J', valor: 'IdPer'},
					{ column: 'K', valor: 'NomPer'},
					{ column: 'L', valor: 'CognomPer'},
					{ column: 'M', valor: 'IdOrg'},
					{ column: 'N', valor: 'NomOrg'}
					]
  }
  read(){
    var workbook = XLSX.readFile(this.file);
    var first_sheet_name = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[first_sheet_name];
    var fila = 3;
    var actualCell;
    var that = this;

    for (actualCell = worksheet['A'+fila]; actualCell; ){
      var dades = [];
			for(let i=0; i < this.columnes.length; i++){
				let dada = this.columnes[i].valor;
				let col = this.columnes[i].column;
				col = col + fila;
				if (worksheet[col])
					dades[dada] = worksheet[col].v;
				else
					dades[dada] = "";
			}
      that.emit('registre', dades);

      fila++;
      actualCell = worksheet['A'+fila];
    }
    if(Sepa.bCheck){
    	that.status = "end";
    }
    if(!Sepa.bCheck){
    	that.status = "process";
    	Sepa.bCheck = true;
    	this.read();
    	
    }
  }
}

class SepaXml {
	constructor() {
		var that = this;
		this.bCheck = false;
		this.iAmount = 0;
		this.iCounter = 0;
		this.sDataCreacio = '';
		this.sDataRemesa = '';
		this.sIdRemesa = '';
		this.bEndPayment = false;
	}
	addPrevInfo() {
		console.log('<?xml version="1.0" encoding="utf-8"?>');
		console.log('<Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.008.001.02" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">');
		console.log('\t<CstmrDrctDbtInitn>');
		console.log('\t\t<GrpHdr>');
		console.log('\t\t\t<MsgId>' + config.TypeDoc + (new Date().getTime()).toString().substring(0,10) + '</MsgId>');
		console.log('\t\t\t<CreDtTm>' + this.sDataCreacio  + '</CreDtTm>');
		console.log('\t\t\t<NbOfTxs>' + this.iCounter  + '</NbOfTxs>');
		console.log('\t\t\t<CtrlSum>' + parseFloat(this.iAmount).toFixed(2)  + '</CtrlSum>');
		console.log('\t\t\t<InitgPty>');
		console.log('\t\t\t\t<Nm>' + config.OrgName + '</Nm>');
		console.log('\t\t\t\t<Id>');
		console.log('\t\t\t\t\t<OrgId>');
		console.log('\t\t\t\t\t\t<Othr>');
		console.log('\t\t\t\t\t\t\t<Id>' + config.OrgId + '</Id>');
		console.log('\t\t\t\t\t\t</Othr>');
		console.log('\t\t\t\t\t</OrgId>');
		console.log('\t\t\t\t</Id>');
		console.log('\t\t\t</InitgPty>');
		console.log('\t\t</GrpHdr>');
		console.log('\t\t<PmtInf>');
		console.log('\t\t\t<PmtInfId>' + this.sIdRemesa + '</PmtInfId>');
		console.log('\t\t\t<PmtMtd>' + 'DD' + '</PmtMtd>');
		console.log('\t\t\t<BtchBookg>' + 'false' + '</BtchBookg>');
		console.log('\t\t\t<NbOfTxs>' + this.iCounter + '</NbOfTxs>');
		console.log('\t\t\t<CtrlSum>' + this.iAmount.toFixed(2) + '</CtrlSum>');
		console.log('\t\t\t<PmtTpInf>');
		console.log('\t\t\t\t<SvcLvl>');
		console.log('\t\t\t\t\t<Cd>' + 'SEPA' + '</Cd>');
		console.log('\t\t\t\t</SvcLvl>');
		console.log('\t\t\t\t<LclInstrm>');
		console.log('\t\t\t\t\t<Cd>' + config.LclInsrt + '</Cd>');
		console.log('\t\t\t\t</LclInstrm>');
		console.log('\t\t\t\t<SeqTp>' + config.SeqTp + '</SeqTp>');
		console.log('\t\t\t</PmtTpInf>');
		console.log('\t\t\t<ReqdColltnDt>' + this.sDataRemesa + '</ReqdColltnDt>');
		console.log('\t\t\t<Cdtr>');
		console.log('\t\t\t\t<Nm>' + config.OrgName + '</Nm>');
		console.log('\t\t\t</Cdtr>');
		console.log('\t\t\t<CdtrAcct>');
		console.log('\t\t\t\t<Id>');
		console.log('\t\t\t\t\t<IBAN>' + config.OrgIban + '</IBAN>');
		console.log('\t\t\t\t</Id>');
		console.log('\t\t\t</CdtrAcct>');
		console.log('\t\t\t<CdtrAgt>');
		console.log('\t\t\t\t<FinInstnId>');
		console.log('\t\t\t\t\t<BIC>' + config.BIC + '</BIC>');
		console.log('\t\t\t\t</FinInstnId>');
		console.log('\t\t\t</CdtrAgt>');
		console.log('\t\t\t<ChrgBr>' + config.Cluausula + '</ChrgBr>');
		console.log('\t\t\t<CdtrSchmeId>');
		console.log('\t\t\t\t<Id>');
		console.log('\t\t\t\t\t<PrvtId>');
		console.log('\t\t\t\t\t\t<Othr>');
		console.log('\t\t\t\t\t\t\t<Id>' + config.OrgId + '</Id>');
		console.log('\t\t\t\t\t\t\t<SchmeNm>');
		console.log('\t\t\t\t\t\t\t\t<Prtry>' + config.Prtry + '</Prtry>');
		console.log('\t\t\t\t\t\t\t</SchmeNm>');
		console.log('\t\t\t\t\t\t</Othr>');
		console.log('\t\t\t\t\t</PrvtId>');
		console.log('\t\t\t\t</Id>');
		console.log('\t\t\t</CdtrSchmeId>');
	}
}
class doMigrarExcel {

  constructor() {
		var that = this
	    this.CC = new RegManager();
	    this.Web = new excelWeb(config.excelWeb);
	    this.count_register = 0;


	    this.CC.on('wrote', function(d){
			that.popRegister();
	    });

	    this.Web.on('registre', function(r){
      		that.pushRegister();
			that.CC.tractarRegistre(r);
		});

		this.Web.on('end', function(){
			that.checkEnd();
		});
  }
  start(){
    this.Web.read();
  }
  pushRegister(){
    this.count_register++;
  }

  popRegister(){
    this.count_register--;
    this.checkEnd();
  }

  checkEnd() {
	if(this.count_register == 0 && this.Web.status == 'process'){
		Sepa.addPrevInfo();
		this.Web.status="end";
		this.start();
	}

	if (this.count_register == 0 && this.Web.status == 'end'){
       	Sepa.iCounter--;
       	if(Sepa.iCounter == 0){
       		Sepa.bEndPayment=true;
       		console.log('\t\t</PmtInf>');
			console.log('\t</CstmrDrctDbtInitn>');
			console.log('</Document>');
       	}
	}
  }
}

var Sepa = new SepaXml();

let myProgram = new doMigrarExcel();
myProgram.start();
