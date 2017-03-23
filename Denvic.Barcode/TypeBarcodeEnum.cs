using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Denvic.Barcode
{
    [ComVisible(true)]
    [Guid("AE2BDAA6-3A5B-4F81-8D9E-CE8903ECB48A")]
    public enum TypeBarcodeEnum
    {
        [Description("EAN_8")]
        EAN_8 = 0,
        [Description("EAN_13")]
        EAN_13 = 1,
        [Description("CODE_39")]
        CODE_39 = 2,
        [Description("CODE_128")]
        CODE_128 = 3,
        [Description("QR_CODE")]
        QR_CODE = 4,
        [Description("DATA_MATRIX")]
        DATA_MATRIX = 5
    }
}
