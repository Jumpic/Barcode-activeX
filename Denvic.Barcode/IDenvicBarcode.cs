using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Denvic.Barcode
{    
    [Guid("9C5BB98E-55B7-4F31-A769-A3DCE3C91D1B")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDenvicBarcode
    {
        string BarcodeValue { get; set; }
        TypeBarcodeEnum BarcodeType { get; set; }
    }
}
