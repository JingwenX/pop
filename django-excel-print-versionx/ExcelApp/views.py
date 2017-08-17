from django.shortcuts import render
from django.template import loader
from django.urls import reverse
from django.http import HttpResponse
from django.http import JsonResponse
import requests
from . import rgen

#Dictionary mapping report id to restful API
d = {'2' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_pop_test/', '18' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_nusery_inspection_test/'}



#index page for gui
def index(request):
	return render(request, 'ExcelApp/index.html')

#main page for filling the form
def details(request):
	return render(request, 'ExcelApp/main.html')

#returns json for testing
def getReport(request):
	#report id
	rid = request.POST['rid']

	#if the id is in the dictionary
	if rid in d:
		response = requests.get(d[rid])

		#rgen.report.formExcel(response.content.decode("utf-8"), rid) # added decode
		rgen.report.formExcel(response.content, rid) 
		file = HttpResponse(response, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		file['Content-Disposition'] = "attachment; filename=test.xlsx"

		#displays the json
		return HttpResponse(file)
	else:
		return HttpResponse('No API in Dictionary')