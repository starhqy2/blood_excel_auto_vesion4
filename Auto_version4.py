# -*- coding:utf-8 -*-
#History:
#2016/8/24 hqy version 4.0
#Any question contact : 302988766@qq.com
from os.path import isdir
from xlwings import *
from os import listdir, mkdir, path,  curdir,  removedirs
import re
import sys
from shutil import rmtree

# set default variables
currentPath = curdir
OriginFileDir = currentPath + "\\00.Origin"
DivExcelResultDir = currentPath + "\\01.DivOrigin"
ExamFiles = currentPath + '\\03.Exam'
PatchResultDir = path.join(currentPath,'02.Result')
TemplatePath = path.join(currentPath,'TemplateFiles')

# Template file in diff type.
Template2Lines = path.join(TemplatePath,'Template2Lines.xlsx')
Template3Lines = path.join(TemplatePath,'Template3Lines.xlsx')

TemplateExam2Lines = path.join(TemplatePath,'TemplateExam2Lines.xlsx')
TemplateExam3Lines = path.join(TemplatePath,'TemplateExam3Lines.xlsx')

TemplateSumUp2Lines = path.join(TemplatePath,'TemplateSumUp2Lines.xlsx')
TemplateSumUp3Lines = path.join(TemplatePath,'TemplateSumUp3Lines.xlsx')

# defult general Template file
Template = Template2Lines
TemplateExam = TemplateExam2Lines
TemplateSumUp = TemplateSumUp2Lines
flag = 2 # indicate 2 line template

# redefine general Template file base on argument.
if len(sys.argv) != 1:
    if sys.argv[1] == str(3):
        Template = Template3Lines
        TemplateExam = TemplateExam3Lines
        TemplateSumUp = TemplateSumUp3Lines
        flag = 3

# divide sheets to different excel file
def divide_excel_sheets():
    # make OriginFiles and DivExcelsResult files
    if not path.isdir(DivExcelResultDir):
        mkdir(DivExcelResultDir)
    if not path.isdir(OriginFileDir):
        mkdir(OriginFileDir)
    # get files list and copy source data to Template
    fileNames = listdir(OriginFileDir)
    for fileName in fileNames:
        filePath = OriginFileDir+'\\'+fileName
        wbOrigin = Book(filePath)
        for eachSheet in wbOrigin.sheets:
            wbTmp = Book()
            wbTmp.sheets[0].name = eachSheet.name
            # Range("A1:M20", wkb=wbTmp).value =Range("A1:M20", wkb = eachSheet).value
            wbTmp.sheets[0].range("A1:M20").value = eachSheet.range("A1:M20").value
            resultPath = DivExcelResultDir+'\\'+path.splitext(fileName)[0]+eachSheet.name.encode('utf-8')+'.xlsx'
            wbTmp.save(resultPath)
            wbTmp.close()
        wbOrigin.close()


# copy data from origin file to result file, generate 02.Result
def gene_result_files():

    # make result dictionary
    if not path.isdir(PatchResultDir):
        mkdir(PatchResultDir)

    # get files list and copy source data to Template
    files = listdir(DivExcelResultDir)
    wbT = Book(Template)

    for file in files:
        if file.endswith(".xlsx"):
            wbA = Book(path.join(DivExcelResultDir, file))
            wbT.sheets[0].range("A26:M34").value = wbA.sheets[0].range("A12:M20").value
            wbA.close()
            wbT.save(r"%s\%s" % (PatchResultDir, re.sub("(.+?).xlsx", r"\1_result.xlsx" , file)))
    wbT.close()
    '''
    #move results to dictionary: result.
    files = listdir(DivExcelResultDir)
    for file in files:
        if file.find("_result") != -1 :
            print(file.find("_result"))
            shutil.move(r"%s\%s" % (currentPath,file) , r"%s\%s" % (DivExcelResultDir,file))
    '''

    # create a Result.xlsx and write summary to it
    files = listdir(PatchResultDir)
    wbR = Book()
    count = 1
    for file in files:

        if file.endswith(".xlsx") and file != "Result.xlsx":
            print(file)
            wbA = Book(r"%s\%s" % (PatchResultDir,file))
            wbR.sheets[0].range("A%s" %  count ).value = file

            wbR.sheets[0].range("B%s:L%s" % (count, count+5)).value = wbA.sheets[0].range("N16:X21").value

            print(wbR.sheets[0].range("B%s:L%s" % (count, count+5)).value)
            wbA.close()
            count += 6

    # format output
    wbR.sheets[0].range("A1:L%s" % str(count-1)).number_format = "0.0" #keep only one decimal digit
    wbR.sheets[0].range("A1:L%s" % str(count-1)).autofit()    #autofit columns

    wbR.save(r"%s\Result.xlsx" % PatchResultDir)
    wbR.close()


# from result generate exam files
def exam_generator(flag):
    # make OriginFiles and DivExcelsResult files
    if not path.isdir(DivExcelResultDir):
        mkdir(DivExcelResultDir)
    if not path.isdir(ExamFiles):
        mkdir(ExamFiles)

    # get files list and copy source data to Template
    DivFileNames = listdir(DivExcelResultDir)
    wbExamTemp = Book(TemplateExam)
    #with open("unCheckFiles.txt", "w") as f:
    for DivFileName in DivFileNames:
        DivFile = path.join(DivExcelResultDir,DivFileName)
        PatchResultPath = path.join(PatchResultDir,path.splitext(DivFileName)[0])+'_result.xlsx'
        try:
            wbOrigin = Book(DivFile)
        except:
            #f.write(DivFile+' Not Found!\n')
            continue
        wbExamTemp.sheets[0].range("A1:M20").value = wbOrigin.sheets[0].range("A1:M20").value
        wbOrigin.close()
        wbPatchResult = Book(PatchResultPath)
        wbExamTemp.sheets[0].range("C22").value = wbPatchResult.sheets[0].range("C4").value # standard curve replace

        if flag == 2:
            # replace titers, 3 lines.
            wbExamTemp.sheets[0].range("C30:L30").value = wbPatchResult.sheets[0].range("O16:X16").value
            wbExamTemp.sheets[0].range("C35:L35").value = wbPatchResult.sheets[0].range("O18:X18").value
            wbExamTemp.sheets[0].range("C40:L40").value = wbPatchResult.sheets[0].range("O20:X20").value
        # copy Template arguments.
        wbExamTemp.sheets[0].range("N22:X40").value = wbPatchResult.sheets[0].range("B4:K25").value
        # replace R2 value.
        wbExamTemp.sheets[0].range("L24").value = wbPatchResult.sheets[0].range("H8").value
        wbExamTemp.save(ExamFiles+'\\'+path.splitext(DivFileName)[0]+'_exam.xlsx')
        wbPatchResult.close()
    wbExamTemp.close()


# sum up final results for 2 lines.
def sum_up(flag):
    wbSumUp = Book(TemplateSumUp) # open SumUp file for input
    exam_file_count = 0
    xList = [x for x in map(chr, range(67, 77))]
    if flag == 3:
        yList = range(27, 34, 6) + range(28, 35, 6) + range(29, 36, 6) + range(30, 37, 6) + range(31, 38, 6)
    else:
        yList = range(27, 38, 5) + range(28, 39, 5) + range(29, 40, 5) + range(30, 41, 5)
    ExamDataList = [x+str(y) for y in yList for x in xList ]

    for ExamFileName in listdir(ExamFiles):
        ExamFile = path.join(ExamFiles, ExamFileName)
        try:
            wbExam = Book(ExamFile)
            exam_file_count += 1
        except:
            continue
        if flag == 3:
            first_end = 22
            xListSumUpGap = (exam_file_count-1)*20
        else:
            first_end = 32
            xListSumUpGap = (exam_file_count-1)*30
        yListSumUp = range(2+xListSumUpGap, first_end+xListSumUpGap)
        if flag == 3:
            xListSumUp = ['A', 'B', 'C', 'D', 'E']
        else:
            xListSumUp = ['A', 'B', 'C', 'D']
        SumUpDataList = [x + str(y) for x in xListSumUp for y in yListSumUp]
        for ExamData, SumUpData in zip(ExamDataList, SumUpDataList):
            wbSumUp.sheets[0].range(SumUpData.encode('utf-8')).value = wbExam.sheets[0].range(ExamData.encode('utf-8')).value
        wbExam.close()
    wbSumUp.sheets[0].range('A1').autofit()
    wbSumUp.save(path.join(currentPath,'SumUp.xlsx'))
    wbSumUp.close()


if __name__ == "__main__":
    for dir in (DivExcelResultDir, PatchResultDir, ExamFiles):
        if isdir(dir):
            rmtree(dir)
    divide_excel_sheets()
    gene_result_files()
    exam_generator(flag)
    sum_up(flag)


