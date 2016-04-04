// Analisi.Dati.Indicedirifrazione.cpp : definisce il punto di ingresso dell'applicazione console.
//

#include "stdafx.h"
#include "BasicExcel.hpp"
#include "ExcelFormat.h"
#include "math.h"

#ifdef _WIN32

#define WIN32_LEAN_AND_MEAN


#include <windows.h>
#include <shellapi.h>
#include <crtdbg.h>

#else // _WIN32


#define	FW_NORMAL	400
#define	FW_BOLD		700

#endif // _WIN32
double sind(double gradi);

double cosd(double gradi);

double convertiSecondiInGradi(double secondi);

double modulo(double val);

double gradiRadianti(double gradi);

double square(double value);

double getAlphaVertice(double riflesso, double minimo);

double getNLambda(double d, double a = 60);

double errorNLambda(double d, double a = 60, double da = 0);

double convertiSecondiInGradi(double secondi) {

	int Gradi = secondi;

	double Gr = Gradi;

	double finale = ((secondi - Gr) / 0.6) + Gr;

	return finale;

};

double gradiRadianti(double gradi) {
	const double M_PI = 4 * atan(1);
	return gradi*M_PI / 180;
};

double sind(double gradi) {
	const double M_PI = 4 * atan(1);
	return sin((gradi)* M_PI / 180);
};

double cosd(double gradi) {
	const double M_PI = 4 * atan(1);
	return cos((gradi)* M_PI / 180);
};



double modulo(double val) {
	if (val < 0.0) {
		return -val;
	}
	else {
		return val;
	};
};

double square(double value) {
	return value*value;
};

double getAlphaVertice(double riflesso, double minimo) {
	return 180 - (minimo + riflesso);
};

double getNLambda(double d, double a) {
	double n;
	n = sind((a + d) / 2) / sind(a / 2);
	return n;
};

double errorNLambda(double d, double a, double da) {
	const double dd = 0.02357;
	return sqrt(square(0.5*(cosd((a+d)/2))*dd*sind(a/2))+square(0.5*da*sind(d/2)/square(sind(a/2))));
};


int main()
{
	YExcel::BasicExcel Vertice("Vertice.xls");

	const double erroreDifferenza = 0.02357;

	const double erroreVertice = 0.03333; //ottenuto tramite legge di propagazione dell'errore

	int number = 0, col = 0, row = 0;

	double alphaVertice;

	YExcel::BasicExcelWorksheet* sheetVertice = Vertice.GetWorksheet(number);

	YExcel::BasicExcelCell* cellVertice = sheetVertice->Cell(row, col);

	YExcel::BasicExcelCell* cellVertice1 = sheetVertice->Cell(row, 0);

	YExcel::BasicExcelCell* cellVertice2 = sheetVertice->Cell(row, 1);

	YExcel::BasicExcelCell* cellVertice3 = sheetVertice->Cell(row, 2);

	YExcel::BasicExcelCell* cellVertice4 = sheetVertice->Cell(row, 3);

	YExcel::BasicExcelCell* cellVertice5 = sheetVertice->Cell(row, 4);

	YExcel::BasicExcelCell* cellVertice6 = sheetVertice->Cell(row, 5);

	/*YExcel::BasicExcelCell* cellVertice7 = sheetVertice->Cell(row, 6);

	YExcel::BasicExcelCell* cellVertice8 = sheetVertice->Cell(row, 7);

	YExcel::BasicExcelCell* cellVertice9 = sheetVertice->Cell(row, 8);

	YExcel::BasicExcelCell* cellVertice10 = sheetVertice->Cell(row, 9);

	YExcel::BasicExcelCell* cellVertice11 = sheetVertice->Cell(row, 10);

	YExcel::BasicExcelCell* cellVertice12 = sheetVertice->Cell(row, 11);*/




	double riflesso = 0, minimo = 0;
	double thetam = 0, thetap = 0, thetapiccolo = 0, thetagrande = 0, riflessom = 0, riflessop = 0;

	thetagrande = cellVertice3->GetDouble();
	thetapiccolo = cellVertice4->GetDouble();
	thetagrande = convertiSecondiInGradi(thetagrande);
	thetapiccolo = convertiSecondiInGradi(thetapiccolo);
	thetap = cellVertice1->GetDouble();
	thetam = cellVertice2->GetDouble();
	thetap = convertiSecondiInGradi(thetap);
	thetam = convertiSecondiInGradi(thetam);
	std::cout << thetap << endl;
	std::cout << thetam << endl;

	minimo = (modulo(thetam - thetapiccolo) + modulo(thetap - thetagrande)) / 2;

	riflessop = cellVertice5->GetDouble();
	riflessom = cellVertice6->GetDouble();
	riflessop = convertiSecondiInGradi(riflessop);
	riflessom = convertiSecondiInGradi(riflessom);

	riflesso = (modulo(360 - thetagrande + riflessop) + modulo(thetapiccolo - riflessom)) / 2;

	alphaVertice = 180 - (minimo + riflesso);

	double aV = alphaVertice;

	std::cout << "row: " << row << ", thetap: " << thetap << ", thetam: " << thetam << ", thetagrande: " << thetagrande << ", thetapiccolo: " << thetapiccolo
		<< ", riflessop: " << riflessop << ", riflessom: " << riflessom << ", minimo: " << minimo << ", errore minimo: " << erroreDifferenza << ", riflesso: " << riflesso << ", errore riflesso: " << erroreDifferenza
		<< ", alphaVertice: " << alphaVertice << ", erroreVertice: " << erroreVertice << endl;
	col = 0;
	sheetVertice->Cell(row, col)->SetDouble(thetap);
	col = 1;
	sheetVertice->Cell(row, col)->SetDouble(thetam);
	col = 2;
	sheetVertice->Cell(row, col)->SetDouble(thetagrande);
	col = 3;
	sheetVertice->Cell(row, col)->SetDouble(thetapiccolo);
	col = 4;
	sheetVertice->Cell(row, col)->SetDouble(riflessop);
	col = 5;
	sheetVertice->Cell(row, col)->SetDouble(riflessom);
	col = 6;
	sheetVertice->Cell(row, col)->SetDouble(minimo);
	col = 7;
	sheetVertice->Cell(row, col)->SetDouble(erroreDifferenza);
	col = 8;
	sheetVertice->Cell(row, col)->SetDouble(riflesso);
	col = 9;
	sheetVertice->Cell(row, col)->SetDouble(erroreDifferenza);
	col = 10;
	sheetVertice->Cell(row, col)->SetDouble(alphaVertice);
	col = 11;
	sheetVertice->Cell(row, col)->SetDouble(erroreVertice);

	row = 1;

	riflesso = 0, minimo = 0;
	thetam = 0, thetap = 0, thetapiccolo = 0, thetagrande = 0, riflessom = 0, riflessop = 0;
	col = 0;
	thetagrande = sheetVertice->Cell(row, col + 2)->GetDouble();
	thetapiccolo = sheetVertice->Cell(row, col + 3)->GetDouble();
	thetagrande = convertiSecondiInGradi(thetagrande);
	thetapiccolo = convertiSecondiInGradi(thetapiccolo);
	thetap = sheetVertice->Cell(row, col)->GetDouble();
	thetam = sheetVertice->Cell(row, col + 1)->GetDouble();
	thetap = convertiSecondiInGradi(thetap);
	thetam = convertiSecondiInGradi(thetam);
	std::cout << thetap << endl;
	std::cout << thetam << endl;

	minimo = (modulo(thetam - thetapiccolo) + modulo(thetap - thetagrande)) / 2;

	riflessop = sheetVertice->Cell(row, col + 4)->GetDouble();
	riflessom = sheetVertice->Cell(row, col + 5)->GetDouble();
	riflessop = convertiSecondiInGradi(riflessop);
	riflessom = convertiSecondiInGradi(riflessom);

	riflesso = (modulo(360 - thetagrande + riflessop) + modulo(thetapiccolo - riflessom)) / 2;

	alphaVertice = 180 - (minimo + riflesso);

	aV = (aV + alphaVertice) / 2;
	double erroreAV = erroreVertice*sqrt(2) / 2;

	std::cout << "row: " << row << ", thetap: " << thetap << ", thetam: " << thetam << ", thetagrande: " << thetagrande << ", thetapiccolo: " << thetapiccolo
		<< ", riflessop: " << riflessop << ", riflessom: " << riflessom << ", minimo: " << minimo << ", errore minimo: " << erroreDifferenza << ", riflesso: " << riflesso << ", errore riflesso: " << erroreDifferenza
		<< ", alphaVertice: " << alphaVertice << ", erroreVertice: " << erroreVertice << endl;
	col = 0;
	sheetVertice->Cell(row, col)->SetDouble(thetap);
	col = 1;
	sheetVertice->Cell(row, col)->SetDouble(thetam);
	col = 2;
	sheetVertice->Cell(row, col)->SetDouble(thetagrande);
	col = 3;
	sheetVertice->Cell(row, col)->SetDouble(thetapiccolo);
	col = 4;
	sheetVertice->Cell(row, col)->SetDouble(riflessop);
	col = 5;
	sheetVertice->Cell(row, col)->SetDouble(riflessom);
	col = 6;
	sheetVertice->Cell(row, col)->SetDouble(minimo);
	col = 7;
	sheetVertice->Cell(row, col)->SetDouble(erroreDifferenza);
	col = 8;
	sheetVertice->Cell(row, col)->SetDouble(riflesso);
	col = 9;
	sheetVertice->Cell(row, col)->SetDouble(erroreDifferenza);
	col = 10;
	sheetVertice->Cell(row, col)->SetDouble(alphaVertice);
	col = 11;
	sheetVertice->Cell(row, col)->SetDouble(erroreVertice);
	col = 13;
	sheetVertice->Cell(row, col)->SetDouble(aV);
	col = 14;
	sheetVertice->Cell(row, col)->SetDouble(erroreAV);

	Vertice.SaveAs("VerticeOutput.xls");

	YExcel::BasicExcel n("n.xls");

	YExcel::BasicExcelWorksheet* nSheet = n.GetWorksheet(0);

	int r, c;
for (r = 0; r < 5; r++) {
	minimo = 0;
	thetam = 0, thetap = 0, thetapiccolo = 0, thetagrande = 0;
	c = 0;
	thetagrande = nSheet->Cell(r, c + 3)->GetDouble();
	thetapiccolo = nSheet->Cell(r, c + 4)->GetDouble();
	thetagrande = convertiSecondiInGradi(thetagrande);
	thetapiccolo = convertiSecondiInGradi(thetapiccolo);
	thetap = nSheet->Cell(r, c + 1)->GetDouble();
	thetam = nSheet->Cell(r, c + 2)->GetDouble();
	thetap = convertiSecondiInGradi(thetap);
	thetam = convertiSecondiInGradi(thetam);
	std::cout << thetap << endl;
	std::cout << thetam << endl;

	minimo = (modulo(thetam - thetapiccolo) + modulo(thetap - thetagrande)) / 2;
	double lambda = nSheet->Cell(r, 0)->GetDouble();
	double n60 = getNLambda(minimo);
	double dN60 = errorNLambda(minimo);

	double nx = getNLambda(minimo, aV);
	double dNx = errorNLambda(minimo, aV, erroreAV);


	c = 0;
	nSheet->Cell(r, c)->SetDouble(lambda);
	c = 1;
	nSheet->Cell(r, c)->SetDouble(thetap);
	c = 2;
	nSheet->Cell(r, c)->SetDouble(thetam);
	c = 3;
	nSheet->Cell(r, c)->SetDouble(thetagrande);
	c = 4;
	nSheet->Cell(r, c)->SetDouble(thetapiccolo);
	c = 5;
	nSheet->Cell(r, c)->SetDouble(minimo);
	c = 6;
	nSheet->Cell(r, c)->SetDouble(erroreDifferenza);
	c = 7;
	nSheet->Cell(r, c)->SetDouble(n60);
	c = 8;
	nSheet->Cell(r, c)->SetDouble(dN60);
	c = 9;
	nSheet->Cell(r, c)->SetDouble(nx);
	c = 10;
	nSheet->Cell(r, c)->SetDouble(dNx);




};

n.SaveAs("nOutput.xls");

	std::cout << "Programma terminato, inserire intero: ";
	int prova;
	std::cin >> prova;





    return 0;
}

