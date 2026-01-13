using ADODB;
using System.Runtime.InteropServices;

namespace TestRecordset
{
    [ComVisible(true)]
    [Guid("7DC55C3B-9A08-4D1D-8022-4578D5C31B44")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRecordsetTlbTest
    {
        [DispId(1)]
        Recordset CreateRecordset();
    }

    [ComVisible(true)]
    [Guid("51B0DAB5-9539-43C4-AA13-F52E346B1539")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("TestRecordset.RecordsetTlbTest")]
    public class RecordsetTlbTest : IRecordsetTlbTest
    {
        public Recordset CreateRecordset()
        {
            return new Recordset();
        }
    }
}

