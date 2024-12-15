import _ctypes
import ctypes
import gc
import os
import pprint
import unittest
import winreg
from unittest import mock

from .win32_com_foreign_func import (
    COINIT_APARTMENTTHREADED,
    E_NOINTERFACE,
    GUID,
    HRESULT,
    POINTER,
    S_OK,
    TRUE,
    COMError,
    CopyComPointer,
    IID_IUnknown,
    byref,
    c_void_p,
    create_guid,
    create_proto_com_method,
    is_equal_guid,
    ole32,
    proto_add_ref,
    proto_query_interface,
    proto_release,
)

oleaut32 = ctypes.oledll.oleaut32
HKCR = winreg.HKEY_CLASSES_ROOT

IN = 1
OUT = 2
RETVAL = 8

COINIT_MULTITHREADED = 0x0

CLSCTX_INPROC_SERVER = 1
CLSCTX_LOCAL_SERVER = 4

REGKIND_DEFAULT = 0
REGKIND_REGISTER = 1
REGKIND_NONE = 2


IID_ITestCtypesComServer = create_guid(
    "{479C32AE-7505-4182-8CF7-F39B7F1A8DE4}"
)
CLSID_TestCtypesComServer = create_guid(
    "{C0A45AA7-4423-4263-9492-4BD6E446823F}"
)
LIBID_TestCtypesComServerLib = create_guid(
    "{4B914909-66E1-47C2-98C1-7A1BD41EA23F}"
)
PROGID_TestCtypesComServer = "TestCtypesComServerLib.TestCtypesComServer.1"
proto_add_one = create_proto_com_method(
    "AddOne", 3, HRESULT, ctypes.c_int, POINTER(ctypes.c_int)
)


def str_from_guid(guid):
    p = ctypes.c_wchar_p()
    ole32.StringFromCLSID(byref(guid), byref(p))
    result = p.value
    ctypes.windll.ole32.CoTaskMemFree(p)
    return result


class IUnknown(c_void_p):
    QueryInterface = proto_query_interface(None, IID_IUnknown)
    AddRef = proto_add_ref(None, IID_IUnknown)
    Release = proto_release(None, IID_IUnknown)


class ITestCtypesComServer(IUnknown):
    AddOne = proto_add_one(
        ((IN, "value"), (OUT | RETVAL, "result")), IID_ITestCtypesComServer
    )


class TestCtypesComServer:
    def AddOne(self, value):
        return value + 1


DIR_NAME = os.path.dirname(__file__)
TLB_FULLPATH = os.path.join(DIR_NAME, "TestCtypesComServer.tlb")
clsid_sub = rf"CLSID\{str_from_guid(CLSID_TestCtypesComServer)}"
inproc_srv_sub = rf"{clsid_sub}\InprocServer32"
full_classname = f"{__name__}.{TestCtypesComServer.__name__}"


def get_inproc_sever_registry_entries():
    return sorted(
        [
            (HKCR, clsid_sub, "", ""),
            (HKCR, rf"{clsid_sub}\ProgID", "", PROGID_TestCtypesComServer),
            (
                HKCR,
                rf"{PROGID_TestCtypesComServer}\CLSID",
                "",
                str_from_guid(CLSID_TestCtypesComServer),
            ),
            (HKCR, inproc_srv_sub, "", _ctypes.__file__),
            (HKCR, inproc_srv_sub, "PythonClass", full_classname),
            (HKCR, inproc_srv_sub, "PythonPath", DIR_NAME),
            (HKCR, inproc_srv_sub, "ThreadingModel", "Both"),
            (
                HKCR,
                rf"{clsid_sub}\Typelib",
                "",
                str_from_guid(LIBID_TestCtypesComServerLib),
            ),
        ]
    )


def register_inproc_server():
    for hkey, subkey, name, value in get_inproc_sever_registry_entries():
        k = winreg.CreateKey(hkey, subkey)
        winreg.SetValueEx(k, name, None, winreg.REG_SZ, str(value))
    ptl = IUnknown()
    oleaut32.LoadTypeLibEx(
        ctypes.c_wchar_p(TLB_FULLPATH), REGKIND_REGISTER, byref(ptl)
    )


def create_instance(clsctx):
    psvr = ITestCtypesComServer()
    # https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance
    ole32.CoCreateInstance(
        byref(CLSID_TestCtypesComServer),
        None,
        clsctx,
        byref(IID_ITestCtypesComServer),
        byref(psvr),
    )
    return psvr


# class IClassFactory(comtypes.IUnknown):
#     _iid_ = comtypes.GUID("{00000001-0000-0000-C000-000000000046}")
#     _methods_ = [
#         comtypes.STDMETHOD(
#             comtypes.HRESULT,
#             "CreateInstance",
#             [
#                 ctypes.POINTER(comtypes.IUnknown),
#                 ctypes.POINTER(comtypes.GUID),
#                 ctypes.POINTER(ctypes.c_void_p),
#             ],
#         ),
#         comtypes.STDMETHOD(comtypes.HRESULT, "LockServer", [ctypes.c_int]),
#     ]


CLASS_E_CLASSNOTAVAILABLE = -2147221231
IID_IClassFactory = create_guid("{00000001-0000-0000-C000-000000000046}")
proto_create_instance = create_proto_com_method(
    "CreateInstance",
    3,
    HRESULT,
    POINTER(IUnknown),
    POINTER(GUID),
    POINTER(c_void_p),
)
proto_lock_server = create_proto_com_method(
    "LockServer", 4, HRESULT, ctypes.c_int
)


class IClassFactory(IUnknown):
    CreateInstance = proto_create_instance(None, IID_IClassFactory)
    LockServer = proto_lock_server(None, IID_IClassFactory)


class TestInprocServer(unittest.TestCase):
    def setUp(self):
        # https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-coinitializeex
        ole32.CoInitializeEx(None, COINIT_MULTITHREADED)
        register_inproc_server()

    def tearDown(self):
        # https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-couninitialize
        ole32.CoUninitialize()
        gc.collect()

    def test_1(self):
        create_instance(CLSCTX_INPROC_SERVER)

    def test_2(self):
        call_args_list = []

        def DllGetClassObject(rclsid, riid, ppv):
            call_args_list.append(
                (
                    str_from_guid(GUID.from_address(rclsid)),
                    str_from_guid(GUID.from_address(riid)),
                    ppv,
                )
            )
            return S_OK

        try:
            with mock.patch.object(
                ctypes, "DllGetClassObject", DllGetClassObject
            ):
                create_instance(CLSCTX_INPROC_SERVER)
        except Exception as e:
            self.fail(f"FAIL:\n{e}\n{pprint.pformat(call_args_list)}")


if __name__ == "__main__":
    unittest.main()
