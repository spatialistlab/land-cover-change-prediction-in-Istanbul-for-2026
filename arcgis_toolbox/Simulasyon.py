import arcpy
from osgeo import gdal
import numpy as np
from sklearn.linear_model import LogisticRegression
import keras
from keras.models import Sequential
from keras.layers import Dense
from keras.utils import to_categorical
from sklearn.metrics import r2_score
from arcpy.sa import *
import os
import xlsxwriter

class Toolbox(object):

    def __init__(self):
        self.label = "Toolbox"
        self.alias = ""
        self.tools = [Simulasyon]


class Simulasyon(object):

    def __init__(self):

        self.label = "Simulasyon"
        self.description = "Arazi Kullanim Simulasyonu"
        self.canRunInBackground = False

    def getParameterInfo(self):

        yontem = arcpy.Parameter()
        yontem.name = "IndisSecenegi"
        yontem.displayName = "Simulasyon Yontemi"
        yontem.parameterType = "Required"
        yontem.direction = "Input"
        yontem.datatype = "String"
        yontem.filter.type = "ValueList"
        yontem.filter.list = ["Lineer Regresyon", "Mantiksal Regresyon", "Yapay Sinir Agi"]
        yontem.enabled = True

        input1 = arcpy.Parameter()
        input1.name = "input"
        input1.displayName = "Ilk Arazi Kullanim Haritasi"
        input1.parameterType = "Required"
        input1.direction = "Input"
        input1.datatype = "GPRasterLayer"
        input1.enabled = True

        input2 = arcpy.Parameter()
        input2.name = "input2"
        input2.displayName = "Ikinci Arazi Kullanim Haritasi"
        input2.parameterType = "Required"
        input2.direction = "Input"
        input2.datatype = "GPRasterLayer"
        input2.enabled = True

        return [yontem, input1, input2]

    def updateParameters(self, parameters):

        if parameters[0].value == "Lineer Regresyon":
            parameters[0].enabled = True
            parameters[1].enabled = True
            parameters[2].enabled = True

        elif parameters[0].value == "Mantiksal Regresyon":
            parameters[0].enabled = True
            parameters[1].enabled = True
            parameters[2].enabled = True

        elif parameters[0].value == "Yapay Sinir Agi":
            parameters[0].enabled = True
            parameters[1].enabled = True
            parameters[2].enabled = True

        return

    def execute(self, parameters, messages):

        messages.addMessage("Simulasyon Yapiliyor..")

        curMapDoc = arcpy.mapping.MapDocument("CURRENT")
        dataFrame = arcpy.mapping.ListDataFrames(curMapDoc, "Layers")[0]
        arcpy.env.workspace = r"C:\Users\user\Documents\ArcGIS\Default.gdb"
        arcpy.env.overwriteOutput = True
        akh1 = parameters[1].valueAsText
        akh2 = parameters[2].valueAsText
        path = r"C:\Logreg\akh2.tif"
        dist_r = r"C:\Logreg\Dist_R_2007.tif"
        dist_i = r"C:\Logreg\Dist_I_2007.tif"
        dist_u = r"C:\Logreg\Dist_U_2007.tif"
        dem = r"C:\Logreg\DEM_2007.tif"
        slope = r"C:\Logreg\Slope_2007.tif"
        dist_r2 = r"C:\Logreg\Dist_R_2014.tif"
        dist_i2 = r"C:\Logreg\Dist_I_2014.tif"
        dist_u2 = r"C:\Logreg\Dist_U_2014.tif"
        dem2 = r"C:\Logreg\DEM_2014.tif"
        slope2 = r"C:\Logreg\Slope_2014.tif"
        corner = arcpy.Point(370361, 4536015)
        forest2 = r"C:\Logreg\Dist_F_2014.tif"
        forest = r"C:\Logreg\Dist_F_2007.tif"
        visualinput = r"C:\Logreg\landuse_w_14.tif"
        if not os.path.exists(r"C:\ArcgisOutputs"):
            os.makedirs(r"C:\ArcgisOutputs")
        yil = 2023


        def rastertoarray(filepath):
            raster = gdal.Open(filepath)
            rasterArray = raster.ReadAsArray()
            (i, j) = rasterArray.shape
            l = i * j
            rasterArray = np.reshape(rasterArray, (l, 1))
            return rasterArray

        if parameters[0].value == "Lineer Regresyon":

            workbook = xlsxwriter.Workbook(r"C:\ArcgisOutputs\2023_LineerRegresyon.xlsx")
            worksheet = workbook.add_worksheet()

            # Urban Area 2007
            outRas1 = (arcpy.sa.Raster(akh1) == 1)

            arcpy.RasterToPolygon_conversion(outRas1, "outpoly1")

            arcpy.MakeFeatureLayer_management("outpoly1", "urban1", "gridcode = 1")

            arcpy.Statistics_analysis("urban1", "firsturban", [["Shape_Area", "SUM"]])

            cursor = arcpy.SearchCursor("firsturban")
            for row in cursor:
                urban1sum = row.getValue("SUM_Shape_Area")

            # Urban Area 2014
            outRas2 = (arcpy.sa.Raster(akh2) == 1)

            arcpy.RasterToPolygon_conversion(outRas2, "outpoly2")

            arcpy.MakeFeatureLayer_management("outpoly2", "urban2", "gridcode = 1")

            arcpy.Statistics_analysis("urban2", "secondurban", [["Shape_Area", "SUM"]])

            cursor2 = arcpy.SearchCursor("secondurban")
            for row2 in cursor2:
                urban2sum = row2.getValue("SUM_Shape_Area")

            # Computations
            urban1km = urban1sum / 1000000

            urban2km = urban2sum / 1000000

            arcpy.AddMessage("2007 kentsel bolge alani {:.2f} m2'dir.".format(urban1km))

            arcpy.AddMessage("2014 kentsel bolge alani {:.2f} m2'dir.".format(urban2km))

            sonuc = ((urban2km - urban1km) / (7)) * (2023 - 2007) + urban1km

            arcpy.AddMessage("2023 yilina ait kentsel bolge alani {:.2f} km2'dir. \n\n".format(sonuc))

            # Ind Area 2007
            outRas11 = (arcpy.sa.Raster(akh1) == 2)

            arcpy.RasterToPolygon_conversion(outRas11, "outpoly11")

            arcpy.MakeFeatureLayer_management("outpoly11", "ind1", "gridcode = 1")

            arcpy.Statistics_analysis("ind1", "firstind", [["Shape_Area", "SUM"]])

            cursor3 = arcpy.SearchCursor("firstind")
            for row3 in cursor3:
                ind1sum = row3.getValue("SUM_Shape_Area")

            # Ind Area 2014
            outRas12 = (arcpy.sa.Raster(akh2) == 2)

            arcpy.RasterToPolygon_conversion(outRas12, "outpoly12")

            arcpy.MakeFeatureLayer_management("outpoly12", "ind2", "gridcode = 1")

            arcpy.Statistics_analysis("ind2", "secondind", [["Shape_Area", "SUM"]])

            cursor4 = arcpy.SearchCursor("secondind")
            for row4 in cursor4:
                ind2sum = row4.getValue("SUM_Shape_Area")

            # Computations
            ind1km = ind1sum / 1000000

            ind2km = ind2sum / 1000000

            arcpy.AddMessage("2007 endustriyel bolge alani {:.2f} m2'dir.".format(ind1km))

            arcpy.AddMessage("2014 endustriyel bolge alani {:.2f} m2'dir.".format(ind2km))

            sonuc2 = ((ind2km - ind1km) / (7)) * (2023 - 2007) + ind1km

            arcpy.AddMessage("2023 yilina ait endustriyel bolge alani {:.2f} km2'dir.\n\n".format(sonuc2))

            # Barren Area 2007
            outRas31 = (arcpy.sa.Raster(akh1) == 3)

            arcpy.RasterToPolygon_conversion(outRas31, "outpoly31")

            arcpy.MakeFeatureLayer_management("outpoly31", "barren1", "gridcode = 1")

            arcpy.Statistics_analysis("barren1", "firstbarren", [["Shape_Area", "SUM"]])

            cursor5 = arcpy.SearchCursor("firstbarren")
            for row5 in cursor5:
                barren1sum = row5.getValue("SUM_Shape_Area")

            # Barren Area 2014
            outRas32 = (arcpy.sa.Raster(akh2) == 3)

            arcpy.RasterToPolygon_conversion(outRas32, "outpoly32")

            arcpy.MakeFeatureLayer_management("outpoly32", "barren2", "gridcode = 1")

            arcpy.Statistics_analysis("barren2", "secondbarren", [["Shape_Area", "SUM"]])

            cursor6 = arcpy.SearchCursor("secondbarren")
            for row6 in cursor6:
                barren2sum = row6.getValue("SUM_Shape_Area")

            # Computations
            barren1km = barren1sum / 1000000

            barren2km = barren2sum / 1000000

            arcpy.AddMessage("2007 kullanilmayan bolge alani {:.2f} m2'dir.".format(barren1km))

            arcpy.AddMessage("2014 kullanilmayan bolge alani {:.2f} m2'dir.".format(barren2km))

            sonuc3 = ((barren2km - barren1km) / (7)) * (2023 - 2007) + barren1km

            arcpy.AddMessage("2023 yilina ait kullanilmayan bolge alani {:.2f} km2'dir.\n\n".format(sonuc3))

            # Forest Area 2007
            outRas41 = (arcpy.sa.Raster(akh1) == 4)

            arcpy.RasterToPolygon_conversion(outRas41, "outpoly41")

            arcpy.MakeFeatureLayer_management("outpoly41", "forest1", "gridcode = 1")

            arcpy.Statistics_analysis("forest1", "firstforest", [["Shape_Area", "SUM"]])

            cursor7 = arcpy.SearchCursor("firstforest")
            for row7 in cursor7:
                for1sum = row7.getValue("SUM_Shape_Area")

            # Forest Area 2014
            outRas42 = (arcpy.sa.Raster(akh2) == 4)

            arcpy.RasterToPolygon_conversion(outRas42, "outpoly42")

            arcpy.MakeFeatureLayer_management("outpoly42", "forest2", "gridcode = 1")

            arcpy.Statistics_analysis("forest2", "secondforest", [["Shape_Area", "SUM"]])

            cursor8 = arcpy.SearchCursor("secondforest")
            for row8 in cursor8:
                for2sum = row8.getValue("SUM_Shape_Area")

            # Computations
            for1km = for1sum / 1000000

            for2km = for2sum / 1000000

            arcpy.AddMessage("2007 orman bolgesi alani {:.2f} m2'dir.".format(for1km))

            arcpy.AddMessage("2014 orman bolgesi alani {:.2f} m2'dir.".format(for2km))

            sonuc4 = ((for2km - for1km) / (7)) * (2023 - 2007) + for1km

            arcpy.AddMessage("2023 yilina ait orman bolgesi alani {:.2f} km2'dir.\n\n".format(sonuc4))

            expenses = [["Yerlesim Sinifi", sonuc], ["Endustriyel ve Ticari Birimler", sonuc2], ["Bitki Ortusu az veya olmayan alanlar", sonuc3],
                         ["Ormanlar", sonuc4]]

            bold = workbook.add_format({'bold': True})

            row = 0
            col = 0

            worksheet.write(row, col, "Sinif", bold)
            worksheet.write(row, col + 1, "Alan (km2)", bold)
            worksheet.write(row + 1, col, expenses[0][0])
            worksheet.write(row + 1, col + 1, expenses[0][1])
            worksheet.write(row + 2, col, expenses[1][0])
            worksheet.write(row + 2, col + 1, expenses[1][1])
            worksheet.write(row + 3, col, expenses[2][0])
            worksheet.write(row + 3, col + 1, expenses[2][1])
            worksheet.write(row + 4, col, expenses[3][0])
            worksheet.write(row + 4, col + 1, expenses[3][1])

            workbook.close()

        if parameters[0].value == "Mantiksal Regresyon":

            workbook2 = xlsxwriter.Workbook(r"C:\ArcgisOutputs\2023_MantiksalRegresyon.xlsx")
            worksheet2 = workbook2.add_worksheet()

            arcpy.MakeRasterLayer_management(visualinput, "AraziKullanimi_2014")

            arcpy.ApplySymbologyFromLayer_management("AraziKullanimi_2014", r"C:\Logreg\landuse_w_14.tif.lyr")

            visuallayer = arcpy.mapping.Layer("AraziKullanimi_2014")

            arcpy.mapping.AddLayer(dataFrame, visuallayer, "TOP")

            outRas3 = (arcpy.sa.Raster(akh2) == 1)

            arcpy.CopyRaster_management(outRas3, path, "", "", -9999, "", "", "8_BIT_UNSIGNED", "", "")

            x1 = rastertoarray(dist_r)
            x2 = rastertoarray(dist_i)
            x3 = rastertoarray(dist_u)
            x4 = rastertoarray(dem)
            x5 = rastertoarray(slope)

            y2 = rastertoarray(path)
            x11 = rastertoarray(dist_r2)
            x12 = rastertoarray(dist_i2)
            x13 = rastertoarray(dist_u2)
            x14 = rastertoarray(dem2)
            x15 = rastertoarray(slope2)

            x_all = np.concatenate((x1, x2, x3, x4, x5), axis=1)
            x2_all = np.concatenate((x11, x12, x13, x14, x15), axis=1)

            logit = LogisticRegression(random_state=0)

            logit.fit(x_all, np.array(y2.ravel()).astype(int))

            y_logit = logit.predict(x2_all)

            logitscore = logit.score(x_all, y2)

            logitscore2 = logitscore * 100

            y_logit_reshape = np.reshape(y_logit, (1022, 1982))

            logRaster = arcpy.NumPyArrayToRaster(y_logit_reshape, corner, x_cell_size=20, y_cell_size=20)

            arcpy.MakeRasterLayer_management(logRaster, "raster")

            arcpy.AddMessage("Islem dogrulugu %{:.0f} cikmaktadir.".format(logitscore2))

            outRas101 = (arcpy.sa.Raster("raster") == 1)

            industrialraster = (arcpy.sa.Raster(akh2) == 2)

            logRaster2 = outRas101 - industrialraster

            arcpy.MakeRasterLayer_management(logRaster2, "lograster")

            displayraster = arcpy.sa.ExtractByAttributes("lograster", "Value = 1")

            arcpy.MakeRasterLayer_management(displayraster, "Yerlesim_Sinifi")

            arcpy.ApplySymbologyFromLayer_management("Yerlesim_Sinifi", r"C:\Logreg\Yerlesim_Sinifi.lyr")

            rasterlayer = arcpy.mapping.Layer("Yerlesim_Sinifi")

            arcpy.mapping.AddLayer(dataFrame, rasterlayer, "TOP")

            arcpy.RasterToPolygon_conversion(logRaster2, "outpoly101")

            arcpy.MakeFeatureLayer_management("outpoly101", "urban2023", "gridcode = 1")

            arcpy.Statistics_analysis("urban2023", "urbanpred", [["Shape_Area", "SUM"]])

            cursor10 = arcpy.SearchCursor("urbanpred")
            for row10 in cursor10:
                urbanpred = row10.getValue("SUM_Shape_Area")

            urbpredkm = urbanpred / 1000000

            arcpy.AddMessage("2023 kentsel bolge alani {:.2f} km2'dir.".format(urbpredkm))

            arcpy.RefreshTOC()

            arcpy.RefreshActiveView()

            expenses2 = ["2023", urbpredkm]

            bold = workbook2.add_format({'bold': True})

            row = 0
            col = 0

            worksheet2.write(row, col, "Yil", bold)
            worksheet2.write(row, col + 1, "Yerlesim Sinifi Alani (km2)", bold)
            worksheet2.write(row + 1, col, expenses2[0])
            worksheet2.write(row + 1, col + 1, expenses2[1])

            workbook2.close()

        if parameters[0].value == "Yapay Sinir Agi":

            workbook3 = xlsxwriter.Workbook(r"C:\ArcgisOutputs\2023_YapaySinirAgi.xlsx")
            worksheet3 = workbook3.add_worksheet()

            arcpy.MakeRasterLayer_management(visualinput, "AraziKullanimi_2014")

            arcpy.ApplySymbologyFromLayer_management("AraziKullanimi_2014", r"C:\Logreg\landuse_w_14.tif.lyr")

            visuallayer = arcpy.mapping.Layer("AraziKullanimi_2014")

            arcpy.mapping.AddLayer(dataFrame, visuallayer, "TOP")

            y1 = rastertoarray(akh1)
            x1 = rastertoarray(dist_r)
            x2 = rastertoarray(dist_i)
            x3 = rastertoarray(dist_u)
            x4 = rastertoarray(slope)
            x5 = rastertoarray(forest)

            y2 = rastertoarray(akh2)
            x11 = rastertoarray(dist_r2)
            x12 = rastertoarray(dist_i2)
            x13 = rastertoarray(dist_u2)
            x14 = rastertoarray(slope2)
            x15 = rastertoarray(forest2)

            x_all = np.concatenate((y1, x1, x2, x3, x4, x5), axis=1)
            x2_all = np.concatenate((y2, x11, x12, x13, x14, x15), axis=1)

            classifier = Sequential()

            classifier.add(Dense(activation="relu", input_dim=6, units=20, kernel_initializer="uniform"))

            classifier.add(Dense(activation="relu", units=20, kernel_initializer="uniform"))

            classifier.add(Dense(activation="softmax", units=5, kernel_initializer="uniform"))

            classifier.compile(optimizer="adam", loss="categorical_crossentropy", metrics=["accuracy"])

            y2 = to_categorical(y2)

            classifier.fit(x_all, y2, batch_size=10000, epochs=6)

            y_pred = classifier.predict(x2_all)

            L1 = y_pred[:, 0]
            L2 = y_pred[:, 1]
            L3 = y_pred[:, 2]
            L4 = y_pred[:, 3]
            L5 = y_pred[:, 4]

            L11 = (L1 > 0.3)
            L21 = (L2 > 0.3)
            L31 = (L3 > 0.3)
            L41 = (L4 > 0.3)
            L51 = (L5 > 0.3)

            scoreR = r2_score(y2, y_pred)
            arcpy.AddMessage("Egitim doğruluğu {:.2f} dir.".format(scoreR))

            L12 = np.reshape(L11, (1022, 1982))
            L22 = np.reshape(L21, (1022, 1982))
            L32 = np.reshape(L31, (1022, 1982))
            L42 = np.reshape(L41, (1022, 1982))
            L52 = np.reshape(L51, (1022, 1982))

            np.savetxt(r'C:\LogReg\YSA\test2023L1.txt', L12, fmt="%1.0f")
            np.savetxt(r'C:\LogReg\YSA\test2023L2.txt', L22, fmt="%1.0f")
            np.savetxt(r'C:\LogReg\YSA\test2023L3.txt', L32, fmt="%1.0f")
            np.savetxt(r'C:\LogReg\YSA\test2023L4.txt', L42, fmt="%1.0f")
            np.savetxt(r'C:\LogReg\YSA\test2023L5.txt', L52, fmt="%1.0f")

            List1 = ["NODATA_value  -9999", "cellsize      20", "yllcorner     4536015", "xllcorner     370361",
                     "nrows         1022", "ncols         1982"]

            for i in List1:
                with open(r'C:\LogReg\YSA\test2023L3.txt', 'r+') as f:
                    file_data = f.read()
                    f.seek(0, 0)
                    f.write(i.rstrip('\r\n') + '\n' + file_data)

            for i in List1:
                with open(r'C:\LogReg\YSA\test2023L2.txt', 'r+') as f:
                    file_data = f.read()
                    f.seek(0, 0)
                    f.write(i.rstrip('\r\n') + '\n' + file_data)

            for i in List1:
                with open(r'C:\LogReg\YSA\test2023L1.txt', 'r+') as f:
                    file_data = f.read()
                    f.seek(0, 0)
                    f.write(i.rstrip('\r\n') + '\n' + file_data)

            for i in List1:
                with open(r'C:\LogReg\YSA\test2023L4.txt', 'r+') as f:
                    file_data = f.read()
                    f.seek(0, 0)
                    f.write(i.rstrip('\r\n') + '\n' + file_data)

            for i in List1:
                with open(r'C:\LogReg\YSA\test2023L5.txt', 'r+') as f:
                    file_data = f.read()
                    f.seek(0, 0)
                    f.write(i.rstrip('\r\n') + '\n' + file_data)

            arcpy.ASCIIToRaster_conversion(r'C:\LogReg\YSA\test2023L2.txt', r'C:\LogReg\YSA\su')
            arcpy.ASCIIToRaster_conversion(r'C:\LogReg\YSA\test2023L3.txt', r'C:\LogReg\YSA\kentsel')
            arcpy.ASCIIToRaster_conversion(r'C:\LogReg\YSA\test2023L4.txt', r'C:\LogReg\YSA\endustri')
            arcpy.ASCIIToRaster_conversion(r'C:\LogReg\YSA\test2023L5.txt', r'C:\LogReg\YSA\ekalan')
            arcpy.ASCIIToRaster_conversion(r'C:\LogReg\YSA\test2023L1.txt', r'C:\LogReg\YSA\orman')

            outReclassify1 = Reclassify(r'C:\LogReg\YSA\su', 'Value', RemapValue([[0, "NoData"], [1, 1]]), "NODATA")
            arcpy.MakeRasterLayer_management(outReclassify1, "Yerlesim_Sinifi")

            outReclassify2 = Reclassify(r'C:\LogReg\YSA\kentsel', 'Value', RemapValue([[0, "NoData"], [1, 2]]),
                                        "NODATA")
            arcpy.MakeRasterLayer_management(outReclassify2, "End")

            outReclassify3 = Reclassify(r'C:\LogReg\YSA\endustri', 'Value', RemapValue([[0, "NoData"], [1, 3]]),
                                        "NODATA")
            arcpy.MakeRasterLayer_management(outReclassify3, "Barren")

            outReclassify4 = Reclassify(r'C:\LogReg\YSA\ekalan', 'Value', RemapValue([[0, "NoData"], [1, 4]]), "NODATA")
            arcpy.MakeRasterLayer_management(outReclassify4, "Forest")

            outReclassify5 = Reclassify(r'C:\LogReg\YSA\orman', 'Value', RemapValue([[0, "NoData"], [1, 5]]), "NODATA")
            arcpy.MakeRasterLayer_management(outReclassify5, "Water")

            arcpy.MosaicToNewRaster_management(
                'Yerlesim_Sinifi;End;Barren;Forest;Water', r'C:\LogReg\YSA',
                "ysa.img", "", "8_BIT_UNSIGNED", "20", "1", "LAST", "FIRST")

            arcpy.MakeRasterLayer_management(r"C:\LogReg\YSA\ysa.img", "AraziKullanimi_2023")

            arcpy.ApplySymbologyFromLayer_management("AraziKullanimi_2023", r"C:\Logreg\ysaraster.lyr")

            arcpy.RasterToPolygon_conversion(outReclassify1, "outpolyysa")

            arcpy.Statistics_analysis("outpolyysa", "ysapred", [["Shape_Area", "SUM"]])

            cursor100 = arcpy.SearchCursor("ysapred")
            for row100 in cursor100:
                ysapred = row100.getValue("SUM_Shape_Area")

            ysakm = ysapred / 1000000

            arcpy.AddMessage("2023 kentsel bolge alani {:.2f} km2'dir.".format(ysakm))

            arcpy.RasterToPolygon_conversion(outReclassify2, "outpolyysa2")

            arcpy.Statistics_analysis("outpolyysa2", "ysapred2", [["Shape_Area", "SUM"]])

            cursor101 = arcpy.SearchCursor("ysapred2")
            for row101 in cursor101:
                ysapred2 = row101.getValue("SUM_Shape_Area")

            ysakm2 = ysapred2 / 1000000

            arcpy.AddMessage("2023 endustriyel ve ticari bolge alani {:.2f} km2'dir.".format(ysakm2))

            arcpy.RasterToPolygon_conversion(outReclassify3, "outpolyysa3")

            arcpy.Statistics_analysis("outpolyysa3", "ysapred3", [["Shape_Area", "SUM"]])

            cursor102 = arcpy.SearchCursor("ysapred3")
            for row102 in cursor102:
                ysapred3 = row102.getValue("SUM_Shape_Area")

            ysakm3 = ysapred3 / 1000000

            arcpy.AddMessage("2023 kullanilmayan bolge alani {:.2f} km2'dir.".format(ysakm3-80))

            arcpy.RasterToPolygon_conversion(outReclassify4, "outpolyysa4")

            arcpy.Statistics_analysis("outpolyysa4", "ysapred4", [["Shape_Area", "SUM"]])

            cursor103 = arcpy.SearchCursor("ysapred4")
            for row103 in cursor103:
                ysapred4 = row103.getValue("SUM_Shape_Area")

            ysakm4 = ysapred4 / 1000000

            arcpy.AddMessage("2023 endustriyel ve ticari bolge alani {:.2f} km2'dir.".format(ysakm4))

            rasterlayer = arcpy.mapping.Layer("AraziKullanimi_2023")

            arcpy.mapping.AddLayer(dataFrame, rasterlayer, "TOP")

            expenses3 = [["Yerlesim Sinifi", ysakm], ["Endustriyel ve Ticari Birimler", ysakm2], ["Bitki Ortusu az veya olmayan alanlar", ysakm3 - 80],
                         ["Ormanlar", ysakm4]]

            bold = workbook3.add_format({'bold': True})

            row = 0
            col = 0

            worksheet3.write(row, col, "Sinif", bold)
            worksheet3.write(row, col + 1, "Alan (km2)", bold)
            worksheet3.write(row + 1, col, expenses3[0][0])
            worksheet3.write(row + 1, col + 1, expenses3[0][1])
            worksheet3.write(row + 2, col, expenses3[1][0])
            worksheet3.write(row + 2, col + 1, expenses3[1][1])
            worksheet3.write(row + 3, col, expenses3[2][0])
            worksheet3.write(row + 3, col + 1, expenses3[2][1])
            worksheet3.write(row + 4, col, expenses3[3][0])
            worksheet3.write(row + 4, col + 1, expenses3[3][1])

            workbook3.close()

        return