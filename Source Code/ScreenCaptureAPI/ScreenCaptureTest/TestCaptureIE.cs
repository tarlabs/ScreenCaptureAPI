using System;
using System.Diagnostics;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Runtime.InteropServices;
using ScreenCaptureAPI;

namespace ScreenCaptureTest
{
    [TestClass]
    public class TestCaptureIE
    {
        [TestMethod]
        public void TestMergeAndCompare()
        {
            ScreenCapture obj_Object = new ScreenCapture();
            const int const_MERGE_IMAGE1TOPIMAGE2 = 3;
            obj_Object.CombineImages("C:\\Temp\\01_006825_AFTER-MT_1.png", "C:\\Temp\\01_006825_AFTER-MT_2.png",
                                     "C:\\Temp\\01_006825_AFTER-MT.png", ScreenCapture.MERGE_Image1LeftImage2, 0);
            obj_Object.CombineImages("C:\\Temp\\01_006825_P-QATHIN_1.png", "C:\\Temp\\01_006825_P-QATHIN_2.png", "C:\\Temp\\01_006825_P-QATHIN.png", ScreenCapture.MERGE_Image1LeftImage2, 0);
            //obj_Object.CompareImages("C:\\Temp\\01_006825_AFTER-MT.png", "C:\\Temp\\01_006825_P-QATHIN.png", true);
            string robj_File1 = "C:\\Temp\\01_006825_AFTER-MT.png";
            string robj_File2 = "C:\\Temp\\01_006825_P-QATHIN.png";
            //Get count of pixels which are different in both images
            int diff = (int) obj_Object.CompareImages(robj_File1, robj_File2, "[PixelDiffCount]");
            Console.Write(diff);
            Debug.Write(diff.ToString());
        }

        [TestMethod]
        public void TestCaptureIEScrollingImage()
        {
            //dynamic oIE;
            //oIE = Activator.CreateInstance(Type.GetTypeFromProgID("InternetExplorer.Application"));
            //oIE.visible = true;
            //oIE.navigate2 ("http://google.com");
            //while (oIE.busy) { };

            ScreenCaptureAPI.ScreenCapture oCap = new ScreenCaptureAPI.ScreenCapture();
            //oCap.CaptureWindow(0x000D0AB6, "c:\\Temp\\IEScrolling.png", true, "", false);
            oCap.CaptureIE(0x000D0AB6, "c:\\Temp\\IEScrolling.png", "", true);
            // oIE.Quit();
            //Marshal.FinalReleaseComObject(oIE);
            //oIE = null;
        }
    }
}
