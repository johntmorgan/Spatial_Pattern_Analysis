##Refactored Python version of spatial pattern program. 
##This version assumes no layer information (e.g. a tissue sample without cortical layers).

####################################################################
##User set variables here

#Directory containing input files
directory = "c:\Documents and Settings\Administrator\Desktop\Spacing UMB Raw"

#Save cell coordinate data from Stereo Investigator as a tab-delimited text file
#There should be a header, leave it in
inputfile = "NoLayer_B4925.txt"

#This is the name of the .xls file that the program will save when finished. 
outputfile = "sp_test_output"

##This variable adjust the cell type that is being compared to
##In this run, neuron = 1, microglia = 3 
cell1 = 1
cell2 = 3

##This variable adjusts the studied distance and interval in the cleaned output
##file - do not exceed the quadratic distance!  Do not exceed excluded distance!
analysis_dist = 100

##This variable sets distance from the layer boundary at which seed cells will be
##excluded from analysis to avoid edge effects
exclude_dist = 100

##This variable adjusts the number of interrogation points over the analysis range
interval_num = 100

##The number of simulations to run; the results are averaged and then the clustering
##values are divided by them to produce a clustering ratio
##In Morgan et al., 2012, I ran 200 simulations/conditin. Experimentation indicated 
##that this number of simulations resulted in very low variability in results.
prun_num = 1

####################################################################
##Program begins here

#Import modules to handle a tab-delimited text file and produce .xls output
import csv
import xlwt
import math
import random
import time
print time.clock()

#load and cleanup file, output is sp_data which has all cells 
#output is [[celltype1, xcoord1, ycoord1],[celltype2, xcoord2, ycoord2], etc]
def loadfile():
    path = directory + "\\" + inputfile
    myfileobj = open(path,"r") 
    csv_read = csv.reader(myfileobj,dialect = csv.excel_tab)
    sp_data = []
    for line in csv_read:
        sp_data.append(line[0:3])
    sp_data = sp_data[1:]
    for cell in sp_data:
        cell[0], cell[1], cell[2] = int(cell[0]), float(cell[1]), float(cell[2])
    return sp_data

#find max and min x and y boundaries
def boundaries(sp_data):
    xmin, ymin, xmax, ymax = sp_data[0][1], sp_data[0][2], sp_data[0][1], sp_data[0][2]
    for cell in sp_data:
        if cell[1] < xmin:
            xmin = cell[1]
        if cell[1] > xmax:
            xmax = cell[1]
        if cell[2] < ymin:
            ymin = cell[2]
        if cell[2] > ymax:
            ymax = cell[2]
    for cell in sp_data:
        xmin_dist = abs(cell[1] - xmin)
        xmax_dist = abs(xmax - cell[1])
        ymin_dist = abs(cell[2] - ymin)
        ymax_dist = abs(ymax - cell[2])
        cell.append(xmin_dist)
        cell.append(xmax_dist)
        cell.append(ymin_dist)
        cell.append(ymax_dist)
    return sp_data, xmin, xmax, ymin, ymax

#Generate clustering values
#This function is the main slowness culprit when launched from Aptana or Eclipse/PyDev.
#The "for cell/for compare_cell" loop runs ~1.5 million times per function call.
#On my PC, with input case B4925 and celltypes set to 1/3:
#Aptana and Eclipse loop in ~.5 sec, while IDLE loops in 6-7 sec.
#Between loops: Aptana/Eclipse <.0001 sec, IDLE: 6-7 sec.
#Given that I will loop through this 1200-6000 times/case doing various calcs, that's huge!
def cluster(sp_data, cell1, cell2):
    print "cluster in: " + str(time.clock())
    raw_cluster = []
    for unused in range(0, analysis_dist + 1):
        raw_cluster.append(0.)
    for cell in sp_data:
        if cell[0] == cell1:
            if cell[3] > exclude_dist and cell[4] > exclude_dist and cell[5] > exclude_dist and cell[6] > exclude_dist:
                #setting these variables here shaves ~7-8% off runtime
                xloc = cell[1]
                yloc = cell[2]
                for compare_cell in sp_data:
                    if compare_cell[0] == cell2:
                        dist = math.sqrt((xloc - compare_cell[1])**2 + (yloc - compare_cell[2])**2)
                        if dist > 0 and dist <= analysis_dist:
                            array_target = int(math.ceil(dist * analysis_dist / interval_num))
                            for insert in range (array_target, analysis_dist + 1):
                                raw_cluster[insert] += 1
    print "cluster out: " + str(time.clock())
    return raw_cluster

#average together the results of the two runs (one from the "perspective" of each cell type)
def cluster_average(cluster1, cluster2):
    for interval in range(0, analysis_dist+1):
        cluster1[interval] = (float(cluster1[interval]) + float(cluster2[interval]))/2
    return cluster1

#makes a simulated version of the cell distribution with random locations
def poisson_gen(sp_data, xmin, xmax, ymin, ymax):
    pois_data = []
    for cell in sp_data:
        pois_data.append([cell[0], random.uniform(xmin, xmax), random.uniform(ymin, ymax), cell[3]])
    return pois_data

#modified version of boundaries function so as not to reset boundaries smaller in simulation runs
#there is probably a better way to refactor all of the boundaries functions
def poisson_boundaries(pois_data, xmin, xmax, ymin, ymax):
    for cell in pois_data:
        xmin_dist = abs(cell[1] - xmin)
        xmax_dist = abs(xmax - cell[1])
        ymin_dist = abs(cell[2] - ymin)
        ymax_dist = abs(ymax - cell[2])
        cell.append(xmin_dist)
        cell.append(xmax_dist)
        cell.append(ymin_dist)
        cell.append(ymax_dist)
    return pois_data

#this is the main function that runs simulations of cellular location
def poisson_iterate(prun_num, sp_data_mod, cell1, cell2, xmin, xmax, ymin, ymax):
    poisson_track = []
    for unused in range(0, analysis_dist + 1):
        poisson_track.append(0)
    for runcount in range(0, prun_num):
        print runcount + 1
        poisson_raw = poisson_boundaries(poisson_gen(sp_data_mod, xmin, xmax, ymin, ymax), xmin, xmax, ymin, ymax)
        if cell1 == cell2:
            poisson_cluster = cluster(poisson_raw, cell1, cell1)
        else:
            poisson_cluster = cluster_average(cluster(poisson_raw, cell1, cell2), cluster(poisson_raw, cell2, cell1))

        for location in range(0,analysis_dist + 1):
            poisson_track[location] = poisson_track[location] + poisson_cluster[location]
    for location in range(0, analysis_dist + 1):
        poisson_track[location] = poisson_track[location] / prun_num
    return poisson_track

#use simulation output to density-correct clustering data
def poisson_correct(raw_cluster, poisson_cluster):
    corrected_output = []
    for unused in range(0, analysis_dist + 1):
        corrected_output.append(0)
    for location in range(0, analysis_dist + 1):
        try:
            corrected_output[location] = raw_cluster[location] / poisson_cluster[location]
        except:
            pass
    return corrected_output

sp_data = loadfile()
sp_data_mod, xmin, xmax, ymin, ymax = boundaries(sp_data)
print sp_data_mod
print xmin, xmax, ymin, ymax

if cell1 == cell2:
    raw_cluster = cluster(sp_data_mod, cell1, cell1)
else:
    raw_cluster = cluster_average(cluster(sp_data_mod, cell1, cell2), cluster(sp_data_mod, cell2, cell1))
print "raw clustering value: "
print raw_cluster

poisson_cluster = poisson_iterate(prun_num, sp_data_mod, cell1, cell2, xmin, xmax, ymin, ymax)
print "poisson clustering value:"
print poisson_cluster

sp_output = poisson_correct(raw_cluster, poisson_cluster)
print "output clustering value: "
print sp_output

print "run time: " + str(time.clock())

#set up worksheet to write to
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Python Sheet 1")

##populate excel worksheet
for location in range(0,analysis_dist+1):
    sheet1.write(0, location, (str(location) + " um"))
    sheet1.write(1, location, sp_output[location])

#save the spreadsheet
savepath = directory + "\\" + outputfile + ".xls"
book.save(savepath)

