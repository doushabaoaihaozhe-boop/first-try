def label_space_describe(dataset_describe,kind):
    '''
    根据dataset_describe文件生成标签描述
    kind: non-code or code
    '''
    if kind == "non-code":
        label_describe = "health"
        for idx, row in dataset_describe.iterrows():
            label_describe += ", "
            label_describe += f"{row.get('location')} fault"
        return label_describe
    elif kind == "code":
        label_describe = "0: healthy state"
        for idx, row in dataset_describe.iterrows():
            label_describe += ", "
            label_describe += f"{int(row['label'])}: {row.get('location')} fault"
        return label_describe
    else:
        raise ValueError(f"Unknown kind: {kind}")

# system输入准备
feature_list, all_feature_dict = feature_list_init(None) #feature_extraction.py
feature_names = ", ".join(list(all_feature_dict.values()))
dataset_describe_dir = dataset_config_yaml["dataset_describe"] 
dataset_describe_sheet_name = "unify" 
dataset_describe = pd.read_excel(dataset_config_yaml["label_describe"],sheet_name=dataset_describe_sheet_name)
fault_modes_num = len(dataset_describe["label"].unique())
label_space_describe_text = label_space_describe(dataset_describe, kind="non-code")

#system提示词
"You are a bearing fault diagnosis expert. \n"
    "Diagnose bearing health state from extracted vibration features with evidence-grounded reasoning. "
    f"Feature order is {feature_names}. The first 12 features are time-domain features; "
    "the last 12 features are frequency-domain features.\n\n"
    f"Probable status includes the following {fault_modes_num} types as follows:"
    f"{label_space_describe_text}\n"
    "The conclusion must include one of the above four status. "
    "Respond with a JSON object: {"label": <int>, "condition": "<str>"}. The extracted features are as follows:...