use windows::core::Interface;
use windows::core::GUID;
use windows::Win32::UI::Shell::PropertiesSystem::PropVariantToStringAlloc;
use windows::core::PWSTR;
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::UI::Shell::PropertiesSystem::IPropertyStore;
use windows::Win32::System::Com::StructuredStorage::STGM_READ;

use windows::Win32::{
    Media::Audio::{
        eAll, eCapture, eConsole, eMultimedia, eRender, ConnectorType, EDataFlow,
        IAudioSessionManager2, IConnector, IDeviceTopology, IMMDevice, IMMDeviceEnumerator,
        IMMEndpoint, IPart, IPartsList, MMDeviceEnumerator, PartType
    },
    System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, CLSCTX_INPROC_SERVER,
        COINIT_APARTMENTTHREADED,
    },
};

fn state_to_string(state: u32) -> Result<String, String> {
    return match state {
        1 => Ok(format!("DEVICE_STATE_ACTIVE '1'")),
        2 => Ok(format!("DEVICE_STATE_DISABLED '2'")),
        4 => Ok(format!("DEVICE_STATE_NOTPRESENT '4'")),
        8 => Ok(format!("DEVICE_STATE_UNPLUGGED '8'")),
        invalid_state => Err(format!("The state '{}' is invalid!", invalid_state)),
    };
}

// pub const eRender: EDataFlow = EDataFlow(0i32);
// pub const eCapture: EDataFlow = EDataFlow(1i32);
// pub const eAll: EDataFlow = EDataFlow(2i32);
// pub const EDataFlow_enum_count: EDataFlow = EDataFlow(3i32);

fn data_flow_to_string(d_flow: EDataFlow) -> String {
    return match d_flow {
        ::windows::Win32::Media::Audio::eRender => format!("eRender"),
        ::windows::Win32::Media::Audio::eCapture => format!("eCapture"),
        ::windows::Win32::Media::Audio::eAll => format!("eAll"),
        ::windows::Win32::Media::Audio::EDataFlow_enum_count => format!("EDataFlow_enum_count"),
        _ => format!("Unknown"),
    };
}

unsafe fn pwstr_to_string(string: PWSTR) -> String {
    let mut end = string.0;
    while *end != 0 {
        end = end.add(1);
    }
    let string_id = String::from_utf16_lossy(std::slice::from_raw_parts(
        string.0,
        end.offset_from(string.0) as _,
    ));
    return string_id;
}

struct audio_node_t {
    state: u32,
    data_flow: EDataFlow,
    node: IPart,
    connector_type: ConnectorType,
}

fn part_type_to_string(part_type: PartType) -> String {
    return match part_type {
        ::windows::Win32::Media::Audio::Connector => format!("Connector"),
        ::windows::Win32::Media::Audio::Subunit => format!("SubUnit"),
        _ => format!("Unknown"),
    };
}

fn connector_type_to_string(connector_type: ConnectorType) -> String {
    return match connector_type {
        ConnectorType::Unknown_Connector => format!("Unknown_Connector"),
        ConnectorType::Physical_Internal => format!("Physical_Internal"),
        ConnectorType::Physical_External => format!("Physical_External"),
        ConnectorType::Software_IO => format!("Software_IO"),
        ConnectorType::Software_Fixed => format!("Software_Fixed"),
        ConnectorType::Network => format!("Network"),
        _ => format!("Unknown"),
    };
}

unsafe fn retrieve_node_details(audio_node: &audio_node_t) {
    let node_name = audio_node
        .node
        .GetName()
        .expect("Couldn't retrieve the name of the node");
    println!(
        "retrieve_node_details: node_name is '{}'",
        pwstr_to_string(node_name)
    );
    let global_id = audio_node
        .node
        .GetGlobalId()
        .expect("Couldn't retrieve the global id of the node");
    println!(
        "retrieve_node_details: global id of node is '{}'",
        pwstr_to_string(global_id)
    );
    let local_id = audio_node
        .node
        .GetLocalId()
        .expect("Couldn't retrieve the local id of the node");
    println!("retrieve_node_details: local id of node is '{}'", local_id);
    let sub_type = audio_node
        .node
        .GetSubType()
        .expect("Couldn't retrieve the subtype of the node");
    println!(
        "retrieve_node_details: local id of node is '{}'",
        sub_type.to_u128()
    );
    let part_type = audio_node
        .node
        .GetPartType()
        .expect("Couldn't retrieve the part type");
    println!(
        "retrieve_node_details: part type is  '{}'",
        part_type_to_string(part_type)
    );
}

unsafe fn enumerate_nodes(audio_node: &audio_node_t, is_last_node: bool) {
    retrieve_node_details(audio_node);

    if is_last_node {
        println!("Incounter the last node !!!!!!");
        return;
    }

    let part_list: IPartsList = if audio_node.data_flow == eRender {
        audio_node
            .node
            .EnumPartsIncoming()
            .expect("Couldn't enumerate the incoming parts for node")
    } else {
        audio_node
            .node
            .EnumPartsOutgoing()
            .expect("Couldn't enumerate the incoming parts for node")
    };

    let number_of_parts = part_list
        .GetCount()
        .expect("Couldn't get the number of parts of the PartList");
    println!("This is the number of parts {}", number_of_parts);

    for part_indx in 0..number_of_parts {
        let child_node: IPart = part_list
            .GetPart(part_indx)
            .expect("Couldn't get the part indx from part_list");
        let node_type = child_node
            .GetPartType()
            .expect("Child node couldn't get part type");
        println!(
            "This is the child part type '{}'",
            part_type_to_string(node_type)
        );

        let last_node = match node_type {
         ::windows::Win32::Media::Audio::Connector => true,
            _ => false,
        };

        let mut connector_type : ConnectorType = ConnectorType::Unknown_Connector;
        if last_node {
            let connector: IConnector = child_node.cast().unwrap();
            connector_type = connector
                .GetType()
                .expect("Couldn't get the type of the last node");
            println!(
                "This is the connector type '{}'",
                connector_type_to_string(connector_type)
            );
        }

        let next_audio_node = audio_node_t {
            state: audio_node.state,
            data_flow: audio_node.data_flow,
            node: child_node,
            connector_type,
        };

        enumerate_nodes(&next_audio_node, last_node);
    }
}

unsafe fn create_audio_node(state: u32, data_flow: EDataFlow, connector: &IConnector) -> audio_node_t {
        let connector_type = connector
            .GetType()
            .expect("Failed to query the type of connector");
        let connected_to: IConnector = connector
            .GetConnectedTo()
            .expect("Failed to query the connected connectors");
        let connected_interface: IPart = connected_to.cast().unwrap();

        return audio_node_t {
            state,
            data_flow,
            node: connected_interface,
            connector_type,
        };
}

unsafe fn create_device_topology(imm_device: &IMMDevice) {
    let mut vector = Vec::new();

    let audio_topology_ref: IDeviceTopology = imm_device
        .Activate(CLSCTX_ALL, std::ptr::null_mut())
        .expect("Activate Failed");

    let connector_count = audio_topology_ref
        .GetConnectorCount()
        .expect("Failed to get the connector count from Audio Topology");

    let imm_device_endpoint: IMMEndpoint = imm_device.cast().unwrap();

    let data_flow = imm_device_endpoint
        .GetDataFlow()
        .expect("Couldn't retrieve the data flow from the imm_device");
    println!("This is the data flow '{}'", data_flow_to_string(data_flow));

    let state = imm_device
        .GetState()
        .expect("Thre was a problem retreiving the enpoint state!");

    println!(
        "This is the state of the current endpoint '{}'",
        state_to_string(state).expect("Couldn't convert state to string")
    );

    for con_index in 0..connector_count {
        let connector: IConnector = audio_topology_ref
            .GetConnector(con_index)
            .expect("Failed to get the connector");

        vector.push(create_audio_node(state, data_flow, &connector));
        retrieve_node_details(&vector[vector.len() - 1]);
        enumerate_nodes(&vector[vector.len() - 1], false);
    }

    println!("This is the list of audio_node_t '{}'", vector.len());
}

fn main() {
    println!("Hello world!!!");
    unsafe {
        CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED).expect("CoInitializeEx Failed");

        // Getting the device enumerator: works
        let imm_device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER)
                .expect("CoCreateInstance Failed");

        let endpoints = imm_device_enumerator
            .EnumAudioEndpoints(eAll, 0x0001)
            .expect("Get enum audio endpoint failed");

        let number_of_endpoints = endpoints
            .GetCount()
            .expect("Failed to get the number of endpoints");

        println!(
            "This are the number of endpoint present in the PC {}",
            number_of_endpoints
        );

        // Endpoint enumeration
        for endpoint_indx in 0..number_of_endpoints {
            let imm_device : IMMDevice = endpoints
                .Item(endpoint_indx)
                .expect("Failed to get the item from the collection of endpoints");
            let endpoint_id = imm_device
                .GetId()
                .expect("Thre was a problem retriving the endpoint id");

            let mut end = endpoint_id.0;
            while *end != 0 {
                end = end.add(1);
            }
            let string_id = String::from_utf16_lossy(std::slice::from_raw_parts(
                endpoint_id.0,
                end.offset_from(endpoint_id.0) as _,
            ));

            let property_store : IPropertyStore = imm_device.OpenPropertyStore(STGM_READ).expect("Couldn't open property store");
            let count : u32 = property_store.GetCount().expect("Couldn't get the number of properties");
            for p_indx in 0..count {
                let prop = property_store.GetAt(p_indx).expect("Couldn't open property at index");
                let value = property_store.GetValue(&prop).expect("Couldn't get the value at index");
                let pwstr_value = PropVariantToStringAlloc(&value).expect("Couldn't convert to PWSTR");
                let string_value = pwstr_to_string(pwstr_value);
                println!("This is the value {}", string_value);
            }

            CoTaskMemFree(endpoint_id.0 as _);

            println!("This is the endpoint id '{}'", string_id);

            create_device_topology(&imm_device);
        }

        // Getting the IMMDevice of the defaultAudioEndpoint: works
        let endpoint = imm_device_enumerator
            .GetDefaultAudioEndpoint(eRender, eConsole)
            .expect("GetDefaultAudioEnpoint Failed");

        // Activating: the target Interface is IAudioSessionManager2: No error!
        let endpoint_id = endpoint
            .GetId()
            .expect("Thre was a problem retriving the endpoint id");

        let mut end = endpoint_id.0;
        while *end != 0 {
            end = end.add(1);
        }
        let string_id = String::from_utf16_lossy(std::slice::from_raw_parts(
            endpoint_id.0,
            end.offset_from(endpoint_id.0) as _,
        ));

        CoTaskMemFree(endpoint_id.0 as _);

        println!("This is the endpoint id '{}'", string_id);

        let state = endpoint
            .GetState()
            .expect("Thre was a problem retreiving the enpoint state!");

        println!(
            "This is the state of the current endpoint '{}'",
            state_to_string(state).expect("Couldn't convert state to string")
        );
    }
}
