#include <windows.h>
#include <stdio.h>

#include <mmdeviceapi.h>
#include <devicetopology.h>

HRESULT WalkTreeBackwardsFromPart(IPart *pPart, int iTabLevel = 0);
HRESULT DisplayVolume(IAudioVolumeLevel *pVolume, int iTabLevel);
HRESULT DisplayMute(IAudioMute *pMute, int iTabLevel);
void Tab(int iTabLevel);

int __cdecl main(void) {
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        printf("Failed CoInitializeEx: hr = 0x%08x\n", hr);
        return __LINE__;
    }

    // get default render endpoint
    IMMDeviceEnumerator *pEnum = NULL;
    hr = CoCreateInstance(
        __uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
        (void**)&pEnum
    );
    if (FAILED(hr)) {
        printf("Couldn't get device enumerator: hr = 0x%08x\n", hr);
        CoUninitialize();
        return __LINE__;
    }
    IMMDevice *pDevice = NULL;
    hr = pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
    if (FAILED(hr)) {
        printf("Couldn't get default render device: hr = 0x%08x\n", hr);
        pEnum->Release();
        CoUninitialize();
        return __LINE__;
    }
    pEnum->Release();

    // get device topology object for that endpoint
    IDeviceTopology *pDT = NULL;
    hr = pDevice->Activate(__uuidof(IDeviceTopology), CLSCTX_ALL, NULL, (void**)&pDT);
    if (FAILED(hr)) {
        printf("Couldn't get device topology object: hr = 0x%08x\n", hr);
        pDevice->Release();
        CoUninitialize();
        return __LINE__;
    }
    pDevice->Release();

    // get the single connector for that endpoint
    IConnector *pConnEndpoint = NULL;
    hr = pDT->GetConnector(0, &pConnEndpoint);
    if (FAILED(hr)) {
        printf("Couldn't get the connector on the endpoint: hr = 0x%08x\n", hr);
        pDT->Release();
        CoUninitialize();
        return __LINE__;
    }
    pDT->Release();

    // get the connector on the device that is
    // connected to
    // the connector on the endpoint
    IConnector *pConnDevice = NULL;
    hr = pConnEndpoint->GetConnectedTo(&pConnDevice);
    if (FAILED(hr)) {
        printf("Couldn't get the connector on the device: hr = 0x%08x\n", hr);
        pConnEndpoint->Release();
        CoUninitialize();
        return __LINE__;
    }
    pConnEndpoint->Release();

    // QI on the device's connector for IPart
    IPart *pPart = NULL;
    hr = pConnDevice->QueryInterface(__uuidof(IPart), (void**)&pPart);
    if (FAILED(hr)) {
        printf("Couldn't get the part: hr = 0x%08x\n", hr);
        pConnDevice->Release();
        CoUninitialize();
        return __LINE__;
    }
    pConnDevice->Release();

    // all the real work is done in this function
    hr = WalkTreeBackwardsFromPart(pPart);
    if (FAILED(hr)) {
        printf("Couldn't walk the tree: hr = 0x%08x\n", hr);
        pPart->Release();
        CoUninitialize();
        return __LINE__;
    }
    pPart->Release();

    CoUninitialize();

    return 0;
}

HRESULT WalkTreeBackwardsFromPart(IPart *pPart, int iTabLevel /* = 0 */) {
    HRESULT hr = S_OK;

    Tab(iTabLevel);
    LPWSTR pwszPartName = NULL;
    hr = pPart->GetName(&pwszPartName);
    if (FAILED(hr)) {
        printf("Could not get part name: hr = 0x%08x", hr);
        return hr;
    }
    printf("Part name: %ws\n", *pwszPartName ? pwszPartName : L"(Unnamed)");
    CoTaskMemFree(pwszPartName);

    // see if this is a volume node part
    IAudioVolumeLevel *pVolume = NULL;
    hr = pPart->Activate(CLSCTX_ALL, __uuidof(IAudioVolumeLevel), (void**)&pVolume);
    if (E_NOINTERFACE == hr) {
        // not a volume node
    } else if (FAILED(hr)) {
        printf("Unexpected failure trying to activate IAudioVolumeLevel: hr = 0x%08x\n", hr);
        return hr;
    } else {
        // it's a volume node...
        hr = DisplayVolume(pVolume, iTabLevel);
        if (FAILED(hr)) {
            printf("DisplayVolume failed: hr = 0x%08x", hr);
            pVolume->Release();
            return hr;
        }

        pVolume->Release();
    }

    // see if this is a mute node part
    IAudioMute *pMute = NULL;
    hr = pPart->Activate(CLSCTX_ALL, __uuidof(IAudioMute), (void**)&pMute);
    if (E_NOINTERFACE == hr) {
        // not a mute node
    } else if (FAILED(hr)) {
        printf("Unexpected failure trying to activate IAudioMute: hr = 0x%08x\n", hr);
        return hr;
    } else {
        // it's a mute node...
        hr = DisplayMute(pMute, iTabLevel);
        if (FAILED(hr)) {
            printf("DisplayMute failed: hr = 0x%08x", hr);
            pMute->Release();
            return hr;
        }

        pMute->Release();
    }

    // get the list of incoming parts
    IPartsList *pIncomingParts = NULL;
    hr = pPart->EnumPartsIncoming(&pIncomingParts);
    if (E_NOTFOUND == hr) {
        // not an error... we've just reached the end of the path
        Tab(iTabLevel);
        printf("No incoming parts at this part\n");
        return S_OK;
    }
    if (FAILED(hr)) {
        printf("Couldn't enum incoming parts: hr = 0x%08x\n", hr);
        return hr;
    }
    UINT nParts = 0;
    hr = pIncomingParts->GetCount(&nParts);
    if (FAILED(hr)) {
        printf("Couldn't get count of incoming parts: hr = 0x%08x\n", hr);
        pIncomingParts->Release();
        return hr;
    }

    // walk the tree on each incoming part recursively
    for (UINT n = 0; n < nParts; n++) {
        IPart *pIncomingPart = NULL;
        hr = pIncomingParts->GetPart(n, &pIncomingPart);
        if (FAILED(hr)) {
            printf("Couldn't get part #%u (0-based) of %u (1-basedSmile hr = 0x%08x\n", n, nParts, hr);
            pIncomingParts->Release();
            return hr;
        }

        hr = WalkTreeBackwardsFromPart(pIncomingPart, iTabLevel + 1);
        if (FAILED(hr)) {
            printf("Couldn't walk tree on part #%u (0-based) of %u (1-basedSmile hr = 0x%08x\n", n, nParts, hr);
            pIncomingPart->Release();
            pIncomingParts->Release();
            return hr;
        }
        pIncomingPart->Release();
    }

    pIncomingParts->Release();

    return S_OK;
}

HRESULT DisplayVolume(IAudioVolumeLevel *pVolume, int iTabLevel) {
    HRESULT hr = S_OK;
    UINT nChannels = 0;

    hr = pVolume->GetChannelCount(&nChannels);

    if (FAILED(hr)) {
        printf("GetChannelCount failed: hr = %08x\n", hr);
        return hr;
    }

    for (UINT n = 0; n < nChannels; n++) {
        float fMinLevelDB, fMaxLevelDB, fStepping, fLevelDB;

        hr = pVolume->GetLevelRange(n, &fMinLevelDB, &fMaxLevelDB, &fStepping);
        if (FAILED(hr)) {
            printf("GetLevelRange failed: hr = 0x%08x\n", hr);
            return hr;
        }

        hr = pVolume->GetLevel(n, &fLevelDB);
        if (FAILED(hr)) {
            printf("GetLevel failed: hr = 0x%08x\n", hr);
            return hr;
        }

        Tab(iTabLevel);
        printf(
            "Channel %u volume is %.3f dB (range is %.3f dB to %.3f dB in increments of %.3f dB)\n",
            n, fLevelDB, fMinLevelDB, fMaxLevelDB, fStepping
        );
    }

    return S_OK;
}

HRESULT DisplayMute(IAudioMute *pMute, int iTabLevel) {
    HRESULT hr = S_OK;
    BOOL bMuted = FALSE;

    hr = pMute->GetMute(&bMuted);

    if (FAILED(hr)) {
        printf("GetMute failed: hr = 0x%08x\n", hr);
        return hr;
    }

    Tab(iTabLevel);
    printf("Mute node: %s\n", bMuted ? "MUTED" : "NOT MUTED");

    return S_OK;
}

void Tab(int iTabLevel) {
    if (0 >= iTabLevel) { return; }
    printf("\t");
    Tab(iTabLevel - 1);
}
