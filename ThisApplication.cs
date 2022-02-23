
#region namespace
using System;
using System.Drawing;
using Microsoft.VisualBasic;
using System.Windows.Forms; //IWin32Window
using System.IO;
using System.Collections; //ArrayList
using System.Collections.Generic; //List<>

using FunctionBay.RecurDyn.ProcessNet;
//For C#
//using FunctionBay.RecurDyn.ProcessNet.Chart;
//using FunctionBay.RecurDyn.ProcessNet.MTT2D;
//using FunctionBay.RecurDyn.ProcessNet.FFlex;
//using FunctionBay.RecurDyn.ProcessNet.RFlex;
//using FunctionBay.RecurDyn.ProcessNet.Tire;

//For VB.NET
//Imports FunctionBay.RecurDyn.ProcessNet
//Imports FunctionBay.RecurDyn.ProcessNet.Chart
//Imports FunctionBay.RecurDyn.ProcessNet.MTT2D
//Imports FunctionBay.RecurDyn.ProcessNet.FFlex
//Imports FunctionBay.RecurDyn.ProcessNet.RFlex
#endregion


namespace LM_Grouser_Profile
{

    using System.IO;
    using System.Text;

    public partial class ThisApplication
    {


        public void Correct_Grouser_Profile()
        {
            StreamReader sr = new StreamReader("E:\\test\\fileFromExcel.csv", Encoding.GetEncoding("euc-kr"));

            int nRow = 0;
            
            List<string> list = new List<string>();

            while (!sr.EndOfStream)
            {
                string s = sr.ReadLine();
                list.Add(s);
                nRow = nRow + 1;
            }
            sr.Close();

            double[,] data = new double[nRow, 3];
            int i = 0;
            foreach (string row in list)
            {
                string[] temp = row.Split(',');
                data[i, 0] = Double.Parse(temp[0]);
                data[i, 1] = Double.Parse(temp[1]);
                double distance = data[i, 0] * data[i, 0] + data[i, 1] * data[i, 1];
                data[i, 2] = distance;
                i = i + 1;
            }

            double min1 = 0.0;
            double min2 = 0.0;
            int min1_index = 0;
            int min2_index = 0;

            double num = 0.0;
            for (int k = 0; k < nRow; k++)
            {
                num = data[k, 2];
                if (k == 0)
                {
                    min1 = num;
                    min2 = num;
                }
                if (num < min1)
                {
                    min1 = num;
                    min1_index = k;
                }
            }

            for (int k = 0; k < nRow; k++)
            {
                num = data[k, 2];
                if (min2 > num && num > min1)
                {
                    min2 = num;
                    min2_index = k;
                }
            }

            int start_number = 0;
            int second_start_number = 0;

            if (data[min1_index, 1] > 0)
            {
                start_number = min1_index;
            }
            else
            {
                start_number = min2_index;
            }

            second_start_number = nRow - start_number;

            List<int> new_order = new List<int>();

            int counter = 0;

            for (int k = 0; k < start_number + 1; k++)
            {
                new_order.Add(start_number + 1 -k);
                counter++;
            }


            for (int p = 0; p < second_start_number - 1; p++)
            {
                new_order.Add(nRow - p);
                counter++;
            }


            List<string> new_profile = new List<string>();
            foreach (int row in new_order)
            {
                string temp = data[row - 1, 0] + "," + data[row - 1 , 1];
                new_profile.Add(temp);

            }

            using (StreamWriter outputFile = new StreamWriter(@"E:\\test\\changedProfile.csv"))
            {
                outputFile.WriteLine("0,0.001");
                foreach (string line in new_profile)
                {
                    outputFile.WriteLine(line);
                    //application.PrintMessage(line);
                }
                outputFile.WriteLine("0,0.001");
                outputFile.Close();
            }
            application.PrintMessage("Mission completed");
        }

        #region VSTA generated code
        private void ThisApplication_Startup(object sender, EventArgs e)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
            MainWindow = new WinWrapper(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
        }

        private void ThisApplication_Shutdown(object sender, EventArgs e)
        {

        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisApplication_Startup);
            this.Shutdown += new System.EventHandler(ThisApplication_Shutdown);
        }
        #endregion

        #region RecurDyn generated code

        #region Common Variables
        static public IApplication application;
        public IModelDocument modelDocument = null;
        public IPlotDocument plotDocument = null;
        public ISubSystem model = null;

        public IReferenceFrame refFrame1 = null;
        public IReferenceFrame refFrame2 = null;

        #endregion

        #region WinForms
        private System.Windows.Forms.IWin32Window MainWindow;

        //If you made your own Form, please Show with argurment, MainWindow
        //MyForm.Show(MainWindow);

        class WinWrapper : System.Windows.Forms.IWin32Window
        {
            public WinWrapper(IntPtr oHandle)
            {
                _oHwnd = oHandle;
            }

            public IntPtr Handle
            {
                get { return _oHwnd; }
            }

            private IntPtr _oHwnd;
        }

        public void CloseWindow()
        {
            System.Windows.Forms.Application.Exit();
        }
        #endregion

        #region Initialize and Dispose By RecurDyn
        [SetUp]
        public void Initialize() //Initialize() will be called automatically before ProcessNet function call.
        {
            application = RecurDynApplication as IApplication;
            modelDocument = application.ActiveModelDocument;
            plotDocument = application.ActivePlotDocument;
            if (modelDocument == null & plotDocument == null)
            {
                application.PrintMessage("No model file");
                modelDocument = application.NewModelDocument("Examples");
            }
            if (modelDocument != null)
            {
                model = modelDocument.Model;
            }
        }

        [TearDown]
        public void Dispose() //Dispose() will be called automatically after ProcessNet function call.
        {
            modelDocument = application.ActiveModelDocument;
            if (modelDocument != null)
            {
                if (modelDocument.Validate() == true)
                {
                    //Redraw() and UpdateDatabaseWindow() can take more time in a heavy model.
                    modelDocument.Redraw();
                    //modelDocument.PostProcess(); //UpdateDatabaseWindow(), SetModified();
                    modelDocument.UpdateDatabaseWindow(); //If you call SetModified(), Animation will be reset.
                    modelDocument.SetModified();
                    modelDocument.SetUndoHistory("ProcessNet");
                }
            }
        }
        #endregion

        #endregion //RecurDyn generated code

    }

    #region Assert Utility Class
    public class Assert
    {
        static public void AreEqualIgnoringCase(string expected, string actual, string message)
        {
            if (0 != String.Compare(expected, actual, true)) //ignoreCase : true
            {
                ThisApplication.application.PrintMessage(message);
            }
        }

        static public void AreEqualIgnoringCase(string expected, string actual)
        {
            String message = "Expected: " + expected + ", Actual: " + actual;
            AreEqualIgnoringCase(expected, actual, message);
        }

        static public void AreEqual(string expected, string actual, string message)
        {
            if (0 != String.Compare(expected, actual, false)) //ignoreCase : false
            {
                ThisApplication.application.PrintMessage(message);
            }
        }

        static public void AreEqual(string expected, string actual)
        {
            String message = "Expected: " + expected + ", Actual: " + actual;
            AreEqual(expected, actual, message);
        }

        static public void AreEqual(double expected, double actual, double delta, string message)
        {
            if (Math.Abs(expected - actual) > delta)
            {
                ThisApplication.application.PrintMessage(message);
            }
        }

        static public void AreEqual(double expected, double actual, double delta)
        {
            String message = "Expected: " + expected + ", Actual: " + actual;
            AreEqual(expected, actual, delta, message);
        }

        static public void AreEqual(double expected, double actual)
        {
            String message = "Expected: " + expected + ", Actual: " + actual;
            AreEqual(expected, actual, 0, message);
        }

        static public void AreEqual(bool expected, bool actual)
        {
            String message = "Expected: " + expected.ToString() + ", Actual: " + actual.ToString();
            AreEqual(expected.ToString(), actual.ToString(), message);
        }

        //static public void AreEqual(Enum expected, Enum actual)
        //{
        //    String message = "Expected: " + expected.ToString() + ", Actual: " + actual.ToString();
        //    AreEqual(expected.ToString(), actual.ToString(), message);
        //}

        static public void AreEqual(object expected, object actual)
        {
            String message = "Expected: " + expected.ToString() + ", Actual: " + actual.ToString();
            AreEqual(expected.ToString(), actual.ToString(), message);
        }

        static public void AreNotEqual(object expected, object actual)
        {
            String message = "Expected: " + expected.ToString() + ", Actual: " + actual.ToString();
            AreNotEqual(expected.ToString(), actual.ToString(), message);
        }

        static public void AreNotEqual(string expected, string actual, string message)
        {
            if (0 == String.Compare(expected, actual, false)) //ignoreCase : false
            {
                ThisApplication.application.PrintMessage(message);
            }
        }

        /// If object is not null Throw the message in output window      
        static public void IsNull(object actual)
        {
            if (actual != null)
                ThisApplication.application.PrintMessage(actual.ToString() + "is not null");
        }

        /// If object is null Throw the message in output window       
        static public void IsNotNull(object actual)
        {
            if (actual == null)
                ThisApplication.application.PrintMessage("Object is null");
        }
    }


    #endregion

    #region Attribute
    //internal use only
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class TestAttribute : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class TearDownAttribute : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class SetUpAttribute : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class CategoryAttribute : Attribute
    {
        public CategoryAttribute(string name) { }
    }

    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class IgnoreAttribute : Attribute
    {
        public IgnoreAttribute(string name) { }
    }
    #endregion

}
