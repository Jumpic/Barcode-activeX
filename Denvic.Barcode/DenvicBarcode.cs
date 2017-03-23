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
    [ComVisible(true)]
    [ProgId("Denvic.Barcode")]
    [Guid("9FAB265B-867F-49D2-9CDC-75E2851B1060")]
    [ClassInterface(ClassInterfaceType.None)]
    public partial class ActiveXBarcode : UserControl, IDenvicBarcode
    {
        private string textBarcode;
        private Bitmap BarcodeImage;
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

        /// <summary>
        /// Инициализация объекта ШК
        /// </summary>
        /// <param name="text"></param>
        private void GenerateBarcode(string text)
        {
            var writer = new BarcodeWriter
            {
                Format = _barcodeFormat,
                Options = new QrCodeEncodingOptions
                {
                    Width = this.Width * 2,
                    Height = this.Height * 2,
                    Margin = 0,
                    ErrorCorrection = ErrorCorrectionLevel.H
                }
            };

            BarcodeImage = writer.Write(text);
        }

        #region registration

        /// <summary>
        /// Регистрация класса в реестре (обязательное условие для отображения в объектах ActiveX например в 1С)
        /// </summary>
        /// <param name="key"></param>
        [ComRegisterFunction()]
        public static void RegisterClass(string key)
        {            
            var skey = new StringBuilder(key);
            skey.Replace(@"HKEY_CLASSES_ROOT\", "");
            using (var regKey = Registry.ClassesRoot.OpenSubKey(skey.ToString(), true))
            {
                using (var ctrl = regKey.CreateSubKey("Control"))
                {
                    ctrl.Close();
                }
                using (var inprocServer32 = regKey.OpenSubKey("InprocServer32", true))
                {
                    inprocServer32.SetValue("CodeBase", Assembly.GetExecutingAssembly().CodeBase);
                    regKey.CreateSubKey("Insertable");
                    inprocServer32.Close();
                }
                regKey.Close();
            }
        }

        /// <summary>
        /// Удаляем информацию из реестра после отмены регистрации
        /// </summary>
        /// <param name="key"></param>
        [ComUnregisterFunction()]        
        public static void UnregisterClass(string key)
        {
            StringBuilder skey = new StringBuilder(key);
            skey.Replace(@"HKEY_CLASSES_ROOT\", "");
            using (var regKey = Registry.ClassesRoot.OpenSubKey(skey.ToString(), true))
            {
                regKey.DeleteSubKey("Control", false);
                regKey.DeleteSubKey("Insertable", false);
                using (var inprocServer32 = regKey.OpenSubKey("InprocServer32", true))
                {
                    regKey.DeleteSubKey("CodeBase", false);
                }
                regKey.Close();
            }
        }
        #endregion

        /// <summary>
        /// Repaint after resizing
        /// Отрисовка после изменения размера
        /// </summary>
        /// <param name="e"></param>
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
            using (var graphic = e.Graphics)
            {
                graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphic.InterpolationMode = InterpolationMode.NearestNeighbor;
                graphic.SmoothingMode = SmoothingMode.None;
                graphic.CompositingQuality = CompositingQuality.Default;
                graphic.CompositingMode = CompositingMode.SourceCopy;

                if (string.IsNullOrWhiteSpace(textBarcode) == false)
                {
                    GenerateBarcode(textBarcode);
                }

                if (BarcodeImage != null)
                {
                    graphic.DrawImage(BarcodeImage, new Rectangle(0, 0, Width, Height), new Rectangle(0, 0, BarcodeImage.Width, BarcodeImage.Height), GraphicsUnit.Pixel);
                }
            }
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // DenvicBarcode
            // 
            Name = "ActiveXControl";
            ResumeLayout(false);

        }        
    }
}
