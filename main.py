import jmespath
import json
import os

in_file = os.path.join(os.path.dirname(__file__), "connectors.json")
out_file = os.path.join(os.path.dirname(__file__), "connectors_scripts.json")
map_file = os.path.join(os.path.dirname(__file__), "map.json")

with open(map_file, 'r') as file:
    mapping = json.load(file)
mapping_sorted = sorted(mapping, key=lambda x: x['Automation'])


with open(map_file, 'w') as file:
    json.dump(mapping_sorted, file, indent=2)


with open(in_file, 'r') as file:
    data = json.load(file)

expression = """
body.value[*].{
"flow_name": properties.displayName,
"Link": id,
"state": properties.state,
"Trigger": properties.definitionSummary.triggers[0].type,
"Trigger2": properties.definitionSummary.triggers[0].kind,
"Action": properties.definitionSummary.actions[],
"Description": properties.definitionSummary.description
}
"""
flows = jmespath.search(expression, data)

filtered_flows = []

for i, flow in enumerate(flows):
    filtered_actions = [action for action in flow['Action'] if action.get('swaggerOperationId') == 'RunScriptProd']
    if filtered_actions:
        filtered_flow = {
            "Name": flow["flow_name"],
            "Link": flow["Link"].replace('/providers/Microsoft.Flow','https://make.powerautomate.com') if flow["Link"] is not None else None,
            "State": "I AM ALIVE" if flow["state"] == "Started" else "I AM DECEASED",
            "Triggered type": flow["Trigger"],
            "Triggered by": flow["Trigger2"],
            "Operation": [action["swaggerOperationId"] for action in filtered_actions][0],
            "Description": flow["Description"],
            "In use with": [action.get("metadata") for action in filtered_actions][0]
        }
        matching_item = next((item for item in mapping_sorted if item["Automation"] == flow["flow_name"]), None)
        if matching_item:
            filtered_flow.update(matching_item)
        filtered_flows.append(filtered_flow)


with open(out_file, 'w') as file:
    json.dump(filtered_flows, file, indent=2)
