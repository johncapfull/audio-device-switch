#include <wchar.h>
#include <windows.h>
#include <Mmdeviceapi.h>
#include "PolicyConfig.h"
#include <Propidl.h>
#include <Functiondiscoverykeys_devpkey.h>

#include <memory>
#include <string>
#include <vector>
#include <iostream>
#include <functional>
#include <algorithm>

namespace {
	template<class T>
	struct Releaser {
		void operator()(T* dev) { dev->Release(); }
	};

	HRESULT SetDefaultAudioPlaybackDevice(LPCWSTR devID)
	{
		IPolicyConfigVista *pPolicyConfig;
		ERole reserved = eConsole;

		HRESULT hr = CoCreateInstance(__uuidof(CPolicyConfigVistaClient),
			NULL, CLSCTX_ALL, __uuidof(IPolicyConfigVista), (LPVOID *)&pPolicyConfig);
		if (SUCCEEDED(hr))
		{
			hr = pPolicyConfig->SetDefaultEndpoint(devID, reserved);
			pPolicyConfig->Release();
		}
		return hr;
	}

	std::wstring getDefaultDeviceId(IMMDeviceEnumerator* pEnum)
	{
		std::wstring id;
		IMMDevice* pDefaultDevice = nullptr;
		LPWSTR defaultID = nullptr;
		HRESULT hr = pEnum->GetDefaultAudioEndpoint(eRender, eMultimedia, &pDefaultDevice);
		if (SUCCEEDED(hr)) {
			std::unique_ptr<IMMDevice, Releaser<IMMDevice>> defaultDevice(pDefaultDevice);
			hr = pDefaultDevice->GetId(&defaultID);
			if (SUCCEEDED(hr)) {
				id = defaultID;
				CoTaskMemFree(defaultID);
			}
		}
		return id;
	}

	void enumDevices(IMMDeviceEnumerator* pEnum, std::function<void(IMMDevice*)> cb)
	{
		IMMDeviceCollection *pDevices;
		HRESULT hr = pEnum->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pDevices);
		if (SUCCEEDED(hr)) {
			std::unique_ptr<IMMDeviceCollection, Releaser<IMMDeviceCollection>> devices(pDevices);
			UINT count;
			devices->GetCount(&count);
			if (SUCCEEDED(hr)) {

				for (size_t i = 0; i < count; i++) {
					IMMDevice *pDevice;
					hr = pDevices->Item(i, &pDevice);
					if (SUCCEEDED(hr)) {
						std::unique_ptr<IMMDevice, Releaser<IMMDevice>> device(pDevice);
						cb(device.get());
					}
				}
			}
		}
	}

	struct DeviceInfo
	{
		std::wstring name;
		std::wstring id;
	};

	DeviceInfo getDeviceInfo(IMMDevice* dev)
	{
		DeviceInfo info;

		LPWSTR wstrID = NULL;
		HRESULT hr = dev->GetId(&wstrID);
		if (SUCCEEDED(hr))
		{
			IPropertyStore *pStore;
			hr = dev->OpenPropertyStore(STGM_READ, &pStore);
			if (SUCCEEDED(hr))
			{
				PROPVARIANT friendlyName;
				PropVariantInit(&friendlyName);
				hr = pStore->GetValue(PKEY_Device_FriendlyName, &friendlyName);
				if (SUCCEEDED(hr))
				{
					info.name = friendlyName.pwszVal;
					info.id = wstrID;

					PropVariantClear(&friendlyName);
				}
				pStore->Release();
			}

			CoTaskMemFree(wstrID);
		}
		return info;
	}

	struct Options {
		int index = -1;
		bool next = false;
		bool special = false;
		bool list = false;
		bool help = false;

		Options(int argc, wchar_t* argv[])
		{
			if (argc == 2) {
				std::wstring argument = argv[1];

				if (argument == L"-n" || argument == L"/n") {
					next = true;
				} else if (argument == L"-l" || argument == L"/l") {
					list = true;
				} else {
					index = std::stoi(argument);
					if (index >= 0) {
						special = true;
					} else {
						help = true;
					}
				}
			} else {
				help = true;
			}
		}
	};

} // ns


int wmain(int argc, wchar_t* argv[])
{
	Options opts(argc, argv);

	if (opts.help) {
		std::cout << 
			"Default sound device switcher.\n" << 
			"-n              select next\n" <<
			"<index>         select device number <index>\n" <<
			"-l              list all devices\n";
		return 0;
	}

	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr)) {
		std::cerr << "Unable to initialize COM";
		return -1;
	}

	IMMDeviceEnumerator *pEnum = NULL;
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);

	if (FAILED(hr)) {
		std::cerr << "Unable to initialize device enumerator";
		return -1;
	}
	std::unique_ptr<IMMDeviceEnumerator, Releaser<IMMDeviceEnumerator>> enumerator(pEnum);

	std::vector<DeviceInfo> infos;
	enumDevices(enumerator.get(), [&infos](IMMDevice* dev)
	{
		infos.push_back(getDeviceInfo(dev));
	});
	
	auto defaultId = getDefaultDeviceId(enumerator.get());

	// Serve request
	if (opts.list) {
		size_t index = 0;
		for (auto& info : infos) {
			bool default = info.id == defaultId;
			std::wcout 
				<< index++ << ": " 
				<< info.name 
				<< (default ? " [default]" : "") << "\n";
		}
	}

	if (opts.special) {
		if (opts.index < 0 || (size_t)opts.index >= infos.size()) {
			std::cerr << "Invalid index, see --help\n";
			return -1;
		}

		SetDefaultAudioPlaybackDevice(infos[opts.index].id.c_str());
	}

	if (opts.next) {
		auto it = std::find_if(std::begin(infos), std::end(infos), [&defaultId](const DeviceInfo &a) {
			return (a.id == defaultId);
		});

		if (it == infos.end() || ++it == infos.end()) {
			it = infos.begin();
		}

		SetDefaultAudioPlaybackDevice(it->id.c_str());
	}
}

