//using Microsoft.VisualStudio.OLE.Interop;
//using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
//using Microsoft.VisualStudio.OLE.Interop;

namespace Denvic.Barcode
{
    //[InterfaceTypeAttribute(1)]
    //[ComConversionLossAttribute]
    //[GuidAttribute("B196B28D-BAB4-101A-B69C-00AA00341D07")]
    //public interface IPropertyPage1
    //{
    //    void Activate(IntPtr HwndPtr, RECT pRect, Int32 bModal);
    //    void Apply();
    //    void Deactivate();
    //    void GetPageInfo(ref PROPPAGEINFO pPageInfo);
    //    void Help([In, MarshalAs(UnmanagedType.LPWStr)] string pszHelpDir);
    //    [PreserveSig]
    //    int IsPageDirty();
    //    void Move(ref RECT pRect);
    //    void SetObjects(uint cObject, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.IUnknown, SizeParamIndex = 0)] object[] pPunk);
    //    void SetPageSite(IPropertyPageSite pPageSite);
    //    void Show(uint mCmdShow);
    //    [PreserveSig]
    //    int ranslateAccelerator(ref Message pMsg);
    //}

    //public class ExamplePropertyPage : IPropertyPage
    //{
    //    PropertyGrid pGrid;

    //    //[DllImport("user32.dll", SetLastError = true)]
    //    //private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    //    public void Activate(IntPtr hWndParent, RECT[] pRect, int bModal)
    //    {
    //        pGrid = new PropertyGrid();
    //        pGrid.CreateControl();
    //        SetParent(pGrid.Handle, hWndParent);
    //    }

    //    public int Apply()
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public void Deactivate()
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public void GetPageInfo(PROPPAGEINFO[] pPageInfo)
    //    {            
    //        pPageInfo[0].cb = (UInt32)Marshal.SizeOf(typeof(PROPPAGEINFO));
    //        pPageInfo[0].pszTitle = "TextBarcode";
    //    }

    //    public void Help(string pszHelpDir)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public int IsPageDirty()
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public void Move(RECT[] pRect)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public void SetObjects(uint cObjects, object[] ppunk)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public void SetPageSite(IPropertyPageSite pPageSite)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public void Show(uint nCmdShow)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public int TranslateAccelerator(MSG[] pMsg)
    //    {
    //        throw new NotImplementedException();
    //    }
    //}


    [ComVisible(true)]
    [ProgId("Denvic.Barcode")]
    [Guid("9FAB265B-867F-49D2-9CDC-75E2851B1060")]
    [ClassInterface(ClassInterfaceType.None)]
    //[ComDefaultInterface(typeof(IDenvicBarcode))]
    //[ComSourceInterfaces(typeof(UserControlEvents))]
    //[System.ComponentModel.EditorBrowsable(EditorBrowsableState.Never)]
    public partial class ActiveXBarcode : UserControl, IDenvicBarcode//, ISpecifyPropertyPages//, IObjectSafety
    {
        //private PictureBox background;
        private string textBarcode;
        private Bitmap BarcodeImage;
        //private IContainer components;
        private BarcodeFormat _barcodeFormat = BarcodeFormat.QR_CODE;

        public ActiveXBarcode()
        {
            BackColor = SystemColors.Window;
            textBarcode = "271431";
            GenerateBarcode(textBarcode);
        }

        public string BarcodeValue
        {
            get
            {
                return textBarcode;
            }

            set
            {
                textBarcode = value;
                Invalidate();
                
            }
        }

        [DisplayName("Тип штрих-кода")]
        [Description("Тип штрих-кода")]
        //[TypeConverter(typeof(EnumTypeConverter))]
        public TypeBarcodeEnum BarcodeType
        {
            get
            {
                return GetTypeBarcode(_barcodeFormat);
            }

            set
            {
                _barcodeFormat = GetFormatBarcode(value);

                if (_barcodeFormat == BarcodeFormat.EAN_13 && textBarcode.Length != 12)
                {
                    textBarcode = "123456789012";
                }
                else if(_barcodeFormat == BarcodeFormat.EAN_8 && textBarcode.Length != 7)
                {
                    textBarcode = "1234567";
                }
                
                Invalidate();
            }
        }

        private TypeBarcodeEnum GetTypeBarcode(BarcodeFormat formatBcode)
        {
            switch (formatBcode)
            {
                case BarcodeFormat.QR_CODE:
                    return TypeBarcodeEnum.QR_CODE;
                case BarcodeFormat.EAN_13:
                    return TypeBarcodeEnum.EAN_13;
                case BarcodeFormat.DATA_MATRIX:
                    return TypeBarcodeEnum.DATA_MATRIX;
                case BarcodeFormat.EAN_8:
                    return TypeBarcodeEnum.EAN_8;
                case BarcodeFormat.CODE_39:
                    return TypeBarcodeEnum.CODE_39;
                case BarcodeFormat.CODE_128:
                    return TypeBarcodeEnum.CODE_128;
                default:
                    return TypeBarcodeEnum.QR_CODE;
            }
        }

        private BarcodeFormat GetFormatBarcode(TypeBarcodeEnum typeBcode)
        {
            switch (typeBcode)
            {
                case TypeBarcodeEnum.QR_CODE:
                    return BarcodeFormat.QR_CODE;
                case TypeBarcodeEnum.EAN_13:
                    return BarcodeFormat.EAN_13;
                case TypeBarcodeEnum.DATA_MATRIX:
                    return BarcodeFormat.DATA_MATRIX;
                case TypeBarcodeEnum.CODE_39:
                    return BarcodeFormat.CODE_39;
                case TypeBarcodeEnum.EAN_8:
                    return BarcodeFormat.EAN_8;
                case TypeBarcodeEnum.CODE_128:
                    return BarcodeFormat.CODE_128;
                default:
                    return BarcodeFormat.QR_CODE;
            }
        }

        private void GenerateBarcode(string text)
        {
            var writer = new BarcodeWriter
            {
                Format = _barcodeFormat,
                Options = new  QrCodeEncodingOptions
                {
                    Width = this.Width * 2,
                    Height = this.Height * 2,
                    Margin = 0,
                    ErrorCorrection = ErrorCorrectionLevel.H
                }
            };

            BarcodeImage = writer.Write(text);

            //Invalidate();

            //ZXing.Common.BitMatrix byteIMGNew = writer.Encode(text);            
            //sbyte[][] imgNew = byteIMGNew.Array;
            //Bitmap bmp1 = new Bitmap(byteIMGNew.Width, byteIMGNew.Height);
            //Graphics g1 = Graphics.FromImage(bmp1);
            //g1.Clear(Color.White);
            //for (int i = 0; i <= imgNew.Length - 1; i++)
            //{
            //    for (int j = 0; j <= imgNew[i].Length - 1; j++)
            //    {
            //        if (imgNew[j][i] == 0)
            //        {
            //            g1.FillRectangle(Brushes.Black, i, j, 1, 1);
            //        }
            //        else
            //        {
            //            g1.FillRectangle(Brushes.White, i, j, 1, 1);
            //        }
            //    }
            //}

            // BackgroundImage = bitmap;            
        }

        [ComRegisterFunction()]
        public static void RegisterClass(string key)
        {
            var skey = new StringBuilder(key);
            skey.Replace(@"HKEY_CLASSES_ROOT\", "");
            var regKey = Registry.ClassesRoot.OpenSubKey(skey.ToString(), true);
            var ctrl = regKey.CreateSubKey("Control");
            ctrl.Close();
            var inprocServer32 = regKey.OpenSubKey("InprocServer32", true);
            inprocServer32.SetValue("CodeBase", Assembly.GetExecutingAssembly().CodeBase);
            regKey.CreateSubKey("Insertable");
            inprocServer32.Close();
            regKey.Close();
        }

        [ComUnregisterFunction()]        
        public static void UnregisterClass(string key)
        {
            StringBuilder skey = new StringBuilder(key);
            skey.Replace(@"HKEY_CLASSES_ROOT\", "");
            RegistryKey regKey = Registry.ClassesRoot.OpenSubKey(skey.ToString(), true);
            regKey.DeleteSubKey("Control", false);
            regKey.DeleteSubKey("Insertable", false);
            RegistryKey inprocServer32 = regKey.OpenSubKey("InprocServer32", true);
            regKey.DeleteSubKey("CodeBase", false);
            regKey.Close();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (!string.IsNullOrWhiteSpace(textBarcode))
            {
               // GenerateBarcode(textBarcode);
            }
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);

            if (!string.IsNullOrWhiteSpace(textBarcode))
            {
                GenerateBarcode(textBarcode);
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            using (var graphic = e.Graphics)
            {

                graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphic.InterpolationMode = InterpolationMode.NearestNeighbor;
                graphic.SmoothingMode = SmoothingMode.None;
                graphic.CompositingQuality = CompositingQuality.Default;
                graphic.CompositingMode = CompositingMode.SourceCopy;

                if (!string.IsNullOrWhiteSpace(textBarcode))
                {
                    GenerateBarcode(textBarcode);
                }

                if (BarcodeImage != null)
                    graphic.DrawImage(BarcodeImage, new Rectangle(0, 0, Width, Height), new Rectangle(0, 0, BarcodeImage.Width, BarcodeImage.Height), GraphicsUnit.Pixel);
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // DenvicBarcode
            // 
            this.Name = "ActiveXControl";
            this.ResumeLayout(false);

        }

        //public void GetPages(CAUUID[] pPages)
        //{

        //pPages[0].cElems = 1;
        //pPages[0].pElems = Marshal.AllocCoTaskMem(16);
        //var g1 = typeof(ExamplePropertyPage).GUID;
        //var b1 = g1.ToByteArray();
        //for (int x = 0; x < 16; x++)
        //{
        //    Marshal.WriteByte(pPages[0].pElems, x, b1[x]);
        //}
        // }

        //public void GetPages(ref Object pPages)
        //{
        //    MessageBox.Show("GetPages");
        //}

        //public enum ObjectSafetyOptions
        //{
        //    INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001,
        //    INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002,
        //    INTERFACE_USES_DISPEX = 0x00000004,
        //    INTERFACE_USES_SECURITY_MANAGER = 0x00000008
        //};

        //public int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions)
        //{
        //    ObjectSafetyOptions m_options = ObjectSafetyOptions.INTERFACESAFE_FOR_UNTRUSTED_CALLER | ObjectSafetyOptions.INTERFACESAFE_FOR_UNTRUSTED_DATA;
        //    pdwSupportedOptions = (int)m_options;
        //    pdwEnabledOptions = (int)m_options;
        //    return 0;
        //}

        //public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        //{
        //    return 0;
        //}
    }
}
