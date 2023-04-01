import ctypes
import enum
import os
import time

import comtypes
import comtypes.client
from PIL import Image
from pystray import Icon, Menu, MenuItem

CLSID_MMDeviceEnumerator = comtypes.GUID('{bcde0395-e52f-467c-8e3d-c4579291692e}')
CLSID_PolicyConfigClient = comtypes.GUID('{870af99c-171d-4f9e-af0d-e63df40c2bc9}')

################################################################################
# wtypes.h
################################################################################
class PROPERTYKEY(ctypes.Structure):
  _fields_ = [
    ('fmtid', comtypes.GUID),
    ('pid', ctypes.wintypes.DWORD)
  ]

class PROPVARIANT(ctypes.Structure):
  _fields_ = [
    ('vt', ctypes.c_int),
    ('pwszVal', ctypes.wintypes.LPWSTR) # TODO: use union, only use for PKEY_Device_FriendlyName
  ]

################################################################################
# FunctionDiscoveryKeys_devpkey.h
################################################################################
PKEY_Device_FriendlyName = PROPERTYKEY(
  comtypes.GUID('{a45c254e-df1c-4efd-8020-67d146a850e0}'),
  14,
)

################################################################################
# propsys.h
################################################################################
STGM_READ	= 0x00000000

class IPropertyStore(comtypes.IUnknown):
  _iid_ = comtypes.GUID('{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}')
  _methods_ = (
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetAt', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetCount', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetValue', 
      (['in'], ctypes.POINTER(PROPERTYKEY), 'key'),
      (['out'], ctypes.POINTER(PROPVARIANT), 'pv'),
    ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'SetValue', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'Commit', ),
  )

################################################################################
# mmdeviceapi.h
################################################################################
DEVICE_STATE_ACTIVE = 0x00000001
DEVICE_STATE_DISABLED = 0x00000002
DEVICE_STATE_NOTPRESENT = 0x00000004
DEVICE_STATE_UNPLUGGED = 0x00000008
DEVICE_STATEMASK_ALL = 0x0000000F

class EDataFlow(enum.IntEnum):
  eRender = 0
  eCapture = 1
  eAll = 2
  EDataFlow_enum_count = 3

class ERole(enum.IntEnum):
  eConsole = 0
  eMultimedia = 1
  eCommunications = 2
  ERole_enum_count = 3

class IMMDevice(comtypes.IUnknown):
  _iid_ = comtypes.GUID('{d666063f-1587-4e43-81f1-b948e807363f}')
  _methods_ = (
    comtypes.COMMETHOD([], comtypes.HRESULT, 'Activate', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'OpenPropertyStore',
      (['in'], ctypes.wintypes.DWORD, 'stgmAccess'),
      (['out'], ctypes.POINTER(ctypes.POINTER(IPropertyStore)), 'ppProperties')
    ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetId',
      (['out'], ctypes.POINTER(ctypes.wintypes.LPWSTR), 'ppstrId')
    ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetState', ),
  )

class IMMDeviceCollection(comtypes.IUnknown):
  _iid_ = comtypes.GUID('{0bd7a1be-7a1a-44db-8397-cc5392387b5e}')
  _methods_ = (
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetCount',
      (['out'], ctypes.POINTER(ctypes.wintypes.UINT), 'pcDevices')
    ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'Item',
      (['in'], ctypes.wintypes.UINT, 'nDevice'),
      (['out'], ctypes.POINTER(ctypes.POINTER(IMMDevice)), 'ppDevice'),
    ),
  )

class IMMDeviceEnumerator(comtypes.IUnknown):
  _iid_ = comtypes.GUID('{a95664d2-9614-4f35-a746-de8db63617e6}')
  _methods_ = (
    comtypes.COMMETHOD([], comtypes.HRESULT, 'EnumAudioEndpoints',
      (['in'], ctypes.c_int, 'dataFlow'),
      (['in'], ctypes.wintypes.DWORD, 'dwStateMask'),
      (['out'], ctypes.POINTER(ctypes.POINTER(IMMDeviceCollection)), 'ppDevices'),
    ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetDefaultAudioEndpoint', 
      (['in'], ctypes.c_int, 'dataFlow'),
      (['in'], ctypes.c_int, 'role'),
      (['out'], ctypes.POINTER(ctypes.POINTER(IMMDevice)), 'ppEndpoint'),
    ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetDevice', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'RegisterEndpointNotificationCallback', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'UnregisterEndpointNotificationCallback', ),
  )

################################################################################
# policyconfig.h
################################################################################
class IPolicyConfig(comtypes.IUnknown):
  _iid_ = comtypes.GUID('{f8679f50-850a-41cf-9c72-430f290290c8}')
  _methods_ = (
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetMixFormat', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetDeviceFormat', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'ResetDeviceFormat', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'SetDeviceFormat', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetProcessingPeriod', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'SetProcessingPeriod', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetShareMode', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'SetShareMode', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'GetPropertyValue', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'SetPropertyValue', ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'SetDefaultEndpoint',
      (['in'], ctypes.wintypes.LPCWSTR, 'deviceId'),
      (['in'], ctypes.c_int, 'role'),
    ),
    comtypes.COMMETHOD([], comtypes.HRESULT, 'SetEndpointVisibility', ),
  )

################################################################################
# api
################################################################################
def get_device_name(device):
  property_storage = device.OpenPropertyStore(STGM_READ)
  name_propvariant = property_storage.GetValue(ctypes.byref(PKEY_Device_FriendlyName))
  return name_propvariant.pwszVal

device_enumerator = comtypes.client.CreateObject(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator)
def get_devices():
  device_collection = device_enumerator.EnumAudioEndpoints(EDataFlow.eRender, DEVICE_STATE_ACTIVE)
  count = device_collection.GetCount()

  dev_list = []
  for i in range(count):
    device = device_collection.Item(i)
    dev_list.append({
      'id': device.GetId(),
      'name': get_device_name(device),
    })

  return dev_list

def get_current_device():
  device = device_enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eConsole)
  return {
    'id': device.GetId(),
    'name': get_device_name(device),
  }

policy_config = comtypes.client.CreateObject(CLSID_PolicyConfigClient, interface=IPolicyConfig)
def change_device(device_id):
  policy_config.SetDefaultEndpoint(device_id, ERole.eConsole)
  policy_config.SetDefaultEndpoint(device_id, ERole.eMultimedia)
  policy_config.SetDefaultEndpoint(device_id, ERole.eCommunications)

VK_CONTROL = 0x11
VK_MENU = 0x12
VK_NUM1 = 0x31
def check_key(*keys):
  value = 0x8000
  for key in keys:
    value &= ctypes.windll.user32.GetAsyncKeyState(key)
  return value & 0x8000 != 0

################################################################################
# main
################################################################################
class DeviceMenuItem(MenuItem):
  def __init__(self, device, callback):
    self.device = device
    super().__init__(
      device['name'],
      lambda: callback(self.device),
      checked=lambda _: get_current_device()['id'] == device['id']
    )

class Tray:
  def __init__(self):
    self.icon = Icon(
      name='audio',
      icon=Image.open(__file__.replace('.py', '.ico')),
      title='audio-switcher',
      menu=self._create_menu()
    )
    Icon.SETUP_THREAD_TIMEOUT = 1

  def _create_menu(self):
    return Menu(
      *[DeviceMenuItem(device, self.select) for device in get_devices()],
      Menu.SEPARATOR,
      MenuItem('Quit', self.quit)
    )

  def select(self, device):
    if get_current_device()['id'] != device['id']:
      change_device(device['id'])
      self.icon.notify(device['name'])
      self.icon.menu = self._create_menu()
      self.icon.update_menu()

  def setup(self, _):
    self.icon.visible = True

    pre_key = None
    while self.icon._running:
      time.sleep(0.5)

      devices = get_devices()
      for i in range(len(devices)):
        if check_key(VK_CONTROL, VK_MENU, VK_NUM1 + i):
          if pre_key == None:
            self.select(devices[i])
          pre_key = VK_NUM1 + i
          break
      else:
        pre_key = None

  def run(self):
    self.icon.run(self.setup)

  def quit(self):
    self.alive = False
    self.icon.stop()

if __name__ == '__main__':
  Tray().run()
