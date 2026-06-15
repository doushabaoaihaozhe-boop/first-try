#本文件基于S1_test_data_make_3.py进行修改
# version:2
# 生成
import numpy as np
import pandas as pd
import typing

def _safe_div(numerator, denominator):
    return 0.0 if denominator == 0 else numerator / denominator

def _frequency_features(signal, sr):
    """
    频域统计量
    返回 dict:
    mean_f, var_f, skew_f, kurt_f,
    gravity_f, std_f, rms_f, avg_f,
    reg_degree, var_param, moment8_f, moment16_f
    """
    x = np.asarray(signal)
    N = len(x)
    if N == 0:
        return {
            "mean_f": 0.0, "var_f": 0.0, "skew_f": 0.0, "kurt_f": 0.0,
            "gravity_f": 0.0, "std_f": 0.0, "rms_f": 0.0, "avg_f": 0.0,
            "reg_degree": 0.0, "var_param": 0.0, "moment8_f": 0.0, "moment16_f": 0.0,
        }

    S = np.abs(np.fft.rfft(x))
    f = np.fft.rfftfreq(N, d=1.0 / sr)
    K = len(S)

    # 频谱幅值统计
    freq_mean = np.mean(S) if K > 0 else 0.0
    centered_S = S - freq_mean
    freq_variance = np.mean(centered_S**2) if K > 0 else 0.0
    freq_skewness = _safe_div(np.mean(centered_S**3), (freq_variance ** 1.5)) if freq_variance > 0 else 0.0
    freq_kurtosis = _safe_div(np.mean(centered_S**4), (freq_variance ** 2)) if freq_variance > 0 else 0.0

    sumS = np.sum(S)
    sum_f2S = np.sum((f**2) * S)
    sum_f4S = np.sum((f**4) * S)

    # gravity frequency
    gravity_freq = _safe_div(np.sum(f * S), sumS)

    # frequency standard deviation
    freq_std = np.sqrt(_safe_div(np.mean(((f - gravity_freq) ** 2) * S), sumS))

    # frequency RMS
    freq_rms = np.sqrt(_safe_div(sum_f2S, sumS))

    # average frequency
    avg_freq = np.sqrt(_safe_div(sum_f4S, sum_f2S))

    # regularity degree
    reg_degree = _safe_div(sum_f2S, np.sqrt(_safe_div(sumS, sum_f4S)))

    # variation parameter
    var_param = _safe_div(freq_std, gravity_freq)

    # 八阶/十六阶矩
    if freq_std > 0:
        moment8_f = _safe_div(np.sum(((f - gravity_freq) ** 3) * S), K * (freq_std ** 3))
        moment16_f = _safe_div(np.sum(((f - gravity_freq) ** 4) * S), K * (freq_std ** 4))
    else:
        moment8_f = 0.0
        moment16_f = 0.0

    return {
        "mean_f": freq_mean,
        "var_f": freq_variance,
        "skew_f": freq_skewness,
        "kurt_f": freq_kurtosis,
        "gravity_f": gravity_freq,
        "std_f": freq_std,
        "rms_f": freq_rms,
        "avg_f": avg_freq,
        "reg_degree": reg_degree,
        "var_param": var_param,
        "moment8_f": moment8_f,
        "moment16_f": moment16_f,
    }

def feature_list_init(features):
    '''
    根据接收到的features参数，返回最终使用的特征列表和**全部**特征的解释字典。
    features参数为None或空列表时，使用全部特征。
    '''
    time_feature_dict = {"mean":"mean value",
                         "std":"standard deviation",
                         "var":"variance",
                         "peak":"peak value",
                         "skew":"skewness",
                         "kurt":"kurtosis",
                         "sqrt_amp":"square root amplitude",
                         "abs_mean":"absolute mean value",
                         "kurt_idx":"kurtosis index",
                         "peak_idx":"peak index",
                         "waveform_idx":"waveform index",
                         "pulse_idx":"pulse index"}

    freq_feature_dict = {"rms_f":"root mean square frequency",
                        "mean_f":"frequency mean value",
                        "var_f":"frequency variance",
                        "skew_f":"frequency skewness",
                        "kurt_f":"frequency kurtosis",
                        "gravity_f":"gravity frequency",
                        "std_f":"frequency standard deviation",
                        "avg_f":"average frequency",
                        "reg_degree":"regularity degree",
                        "var_param":"variation parameter",
                        "moment8_f":"eighth-order moment",
                        "moment16_f":"sixteenth-order moment"}
    feature_list = list(time_feature_dict.keys()) + list(freq_feature_dict.keys())
    feature_desc = {**time_feature_dict, **freq_feature_dict} #合并字典的方法

    selected = list(dict.fromkeys(feature_list if features is None else features))
    unknown = set(selected) - set(feature_list)
    if unknown:
        raise ValueError(f"Unknown features requested: {sorted(unknown)}")
    return selected, feature_desc

def feature_extraction(data, sampling_freq, features=None) -> typing.Tuple[pd.DataFrame, typing.Dict[str, str]]:
    """
    提取指定的特征；features 为空时计算全部特征。
    return: 特征 DataFrame，已计算特征的解释字典
    """
    selected, feature_desc = feature_list_init(features)
    
    data_features = pd.DataFrame(index=data.index)

    if "mean" in selected:
        data_features["mean"] = data.apply(lambda x: np.mean(x), axis=1)
    if "std" in selected:
        data_features["std"] = data.apply(lambda x: np.std(x), axis=1)
    if "var" in selected:
        data_features["var"] = data.var(axis=1)

    need_peak = any(k in selected for k in ["peak", "peak_idx", "pulse_idx"])
    if need_peak:
        peak_series = data.abs().max(axis=1)
        if "peak" in selected:
            data_features["peak"] = peak_series

    if "skew" in selected:
        data_features["skew"] = data.skew(axis=1)
    if "kurt" in selected:
        data_features["kurt"] = data.kurtosis(axis=1)

    if "sqrt_amp" in selected:
        data_features["sqrt_amp"] = data.apply(lambda x: (np.mean(np.sqrt(np.abs(x))))**2, axis=1)
    if "abs_mean" in selected:
        data_features["abs_mean"] = data.abs().mean(axis=1)

    # 依赖关系特征
    need_abs_mean = any(k in selected for k in ["abs_mean", "waveform_idx", "pulse_idx"])
    if need_abs_mean and "abs_mean" not in data_features.columns:
        abs_mean_series = data.abs().mean(axis=1)
        if "abs_mean" in selected:
            data_features["abs_mean"] = abs_mean_series
    elif "abs_mean" in data_features.columns:
        abs_mean_series = data_features["abs_mean"]
    else:
        abs_mean_series = None

    if "kurt_idx" in selected:
        if "kurt" in data_features.columns:
            kurt_series = data_features["kurt"]
        else:
            kurt_series = data.kurtosis(axis=1)
        if "std" in data_features.columns:
            std_series = data_features["std"]
        else:
            std_series = data.apply(lambda x: np.std(x), axis=1)
        data_features["kurt_idx"] = kurt_series / (std_series ** 2)

    if "peak_idx" in selected:
        if "std" in data_features.columns:
            std_series = data_features["std"]
        else:
            std_series = data.apply(lambda x: np.std(x), axis=1)
        if need_peak:
            data_features["peak_idx"] = peak_series / std_series
        else:
            data_features["peak_idx"] = data.abs().max(axis=1) / std_series

    if "waveform_idx" in selected:
        if "std" in data_features.columns:
            std_series = data_features["std"]
        else:
            std_series = data.apply(lambda x: np.std(x), axis=1)
        data_features["waveform_idx"] = std_series / abs_mean_series

    if "pulse_idx" in selected:
        if need_peak:
            data_features["pulse_idx"] = peak_series / abs_mean_series
        else:
            data_features["pulse_idx"] = data.abs().max(axis=1) / abs_mean_series

    extra_freq_keys = {
        "mean_f", "var_f", "skew_f", "kurt_f", "gravity_f", "std_f", "rms_f", "avg_f",
        "reg_degree", "var_param", "moment8_f", "moment16_f"
    }
    if any(k in selected for k in extra_freq_keys):
        freq_df = data.apply(
            lambda row: pd.Series(_frequency_features(row.values, sampling_freq)), axis=1
        )
        for k in selected:
            if k in freq_df.columns:
                data_features[k] = freq_df[k]

    selected_desc = {name: feature_desc[name] for name in selected}
    return data_features, selected_desc