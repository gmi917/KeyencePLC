using System.Runtime.InteropServices;

namespace KvHostLink
{
    // Socket Type
    public enum KHLSockType : byte
    {
        SOCK_TCP = 0,
        SOCK_UDP = 1,
    }

    // KV Mode
    public enum KHLMode : byte
    {
        MODE_RUN = 0x01,
        MODE_PROG = 0x02,
    }

    // KV Model
    public enum KHLModelCode : byte
    {
        KV8000 = 0x17,
        KV7500 = 0x16,
        KV7300 = 0x15,
        KVNC32 = 0x14,
        KVN60 = 0x13,
        KVN40 = 0x12,
        KVN24 = 0x11,
        KVN14 = 0x10,
    }

    // Device Type
    public enum KHLDevType : byte
    {
        DEV_R = 0x00,
        DEV_CR = 0x01,
        DEV_T = 0x02, // timer contract
        DEV_C = 0x03, // counter contract
        DEV_DM = 0x04,
        DEV_CM = 0x05,
        DEV_TM = 0x06,
        DEV_MR = 0x07,
        DEV_LR = 0x08,
        DEV_VB = 0x09,
        DEV_EM = 0x0A,
        DEV_FM = 0x0B,
        DEV_B = 0x0C,
        DEV_W = 0x0D,
        DEV_VM = 0x0E,
        DEV_TC = 0x0F, // timer current val
        DEV_TS = 0x10, // timer setting val
        DEV_CC = 0x11, // counter current val
        DEV_CS = 0x12, // counter setting val
        DEV_AT = 0x13, // digital trimer
        DEV_ZF = 0x14,
        DEV_UG = 0x15, // buffer memory
        DEV_Z = 0x16, // index register

        // XYM represent
        DEV_X = 0x17,
        DEV_Y = 0x18,
        DEV_M = 0x19,
        DEV_L = 0x1A,
        DEV_D = 0x1B,
        DEV_E = 0x1C,
        DEV_F = 0x1D,
    }

    public class KvHostLink 
    {
        // Initialize
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLInit();

        // Connect
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLConnect(string ipAddr, ushort port, uint timeOutMs, KHLSockType sockType, ref int hSock);
        
        // Disconnect
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLDisconnect(int hSock);


        // Confirm connection status
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLIsConnected(int hSock);

        // Get PLC Model Code
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLGetModelCode(int hSock, ref KHLModelCode modeCode);

        // Get PLC Mode
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLGetMode(int hSock, ref KHLMode mode);
        
        // Change PLC Mode
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLChangeMode(int hSock, KHLMode mode);

        // Get CPU Error Code
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLGetErrorCode(int hSock, ref byte errCode, ref byte errLevel);

        // CPU Error Clear
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLErrorClear(int hSock);

        // Read CPU Global Device
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLReadDevicesAsWords(int hSock, KHLDevType devType, uint devTopNo, uint wordNum, byte[] data);

        // Read CPU Global Device
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLReadDevicesAsBits(int hSock, KHLDevType devType, uint devTopNo, ushort bitOffset, uint bitNum, byte[] data);

        // Write CPU Global Device
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLWriteDevicesAsWords(int hSock, KHLDevType devType, uint devTopNo, uint wordNum, byte[] data);

        // Write CPU Global Device
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLWriteDevicesAsBits(int hSock, KHLDevType devType, uint devTopNo, ushort bitOffset, uint bitNum, byte[] data);

        // Read Extension Unit Device
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLReadExtUnitDevicesAsWords(int hSock, byte unitNo, KHLDevType devType, uint devTopNo, uint wordNum, byte[] data);

        // Read Extension Unit Device
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLReadExtUnitDevicesAsBits(int hSock, byte unitNo, KHLDevType devType, uint devTopNo, ushort bitOffset, uint bitNum, byte[] data);

        // Write Extension Unit Device
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLWriteExtUnitDevicesAsWords(int hSock, byte unitNo, KHLDevType devType, uint devTopNo, uint wordNum, byte[] data);

        // Write Extension Unit Device
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLWriteExtUnitDevicesAsBits(int hSock, byte unitNo, KHLDevType devType, uint devTopNo, ushort bitOffset, uint bitNum, byte[] data);


        // Resolve Variable Name
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLResolveVariableName(int hSock, string varName);

        // Read Variable
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLReadVariable(int hSock, string varName, byte[] data, bool isChkCnv);

        // Write Variable
        [DllImport("KvHostLinkBase.dll", CharSet = CharSet.Unicode)]
        public extern static int KHLWriteVariable(int hSock, string varName, byte[] data, bool isChkCnv);

    }
}

