import os
import re

def validate_params(task_id, params, task_metadata, skip_exists=False):
    """
    通用参数验证引擎
    :param task_id: 任务ID
    :param params: 用户提交的参数字典
    :param task_metadata: 任务元数据 (包含 rules)
    :param skip_exists: 是否跳过文件存在性校验 (用于队列任务，因为文件可能由前置任务动态生成)
    :return: (is_valid, error_msg)
    """
    if task_id not in task_metadata:
        return False, f"未定义的任务 ID: {task_id}"
    
    meta = task_metadata[task_id]
    param_meta_list = meta.get("params", [])
    
    for p_meta in param_meta_list:
        p_id = p_meta["id"]
        
        # info 类型的参数纯作展示，不需要校验
        if p_meta.get("type") == "info":
            continue

        rules = p_meta.get("rules", {})
        value = params.get(p_id, "").strip() if isinstance(params.get(p_id), str) else params.get(p_id)
        
        # 1. 必填检查
        if p_meta.get("required", True) and not value and value != 0:
            return False, f"参数 '{p_meta['label']}' 是必填项"
        
        if not value and value != 0:
            continue

        # 2. 存在性与类型检查 (针对路径)
        if not skip_exists and (rules.get("exists") or rules.get("isdir") or rules.get("isfile")):
            if not os.path.exists(str(value)):
                return False, f"路径不存在: {value} (来自 '{p_meta['label']}')"
            
            if rules.get("isdir") and not os.path.isdir(str(value)):
                return False, f"目标必须是文件夹: {value}"
            
            if rules.get("isfile") and not os.path.isfile(str(value)):
                return False, f"目标必须是文件: {value}"

        # 3. 后缀名检查
        if "ext" in rules:
            exts = rules["ext"]
            _, actual_ext = os.path.splitext(str(value).lower())
            if actual_ext not in [e.lower() for e in exts]:
                return False, f"非法文件格式: {value} (仅支持 {', '.join(exts)})"

        # 4. 数值检查
        if p_meta.get("type") == "number":
            try:
                num_val = float(value)
                if "min" in rules and num_val < rules["min"]:
                    return False, f"'{p_meta['label']}' 不能小于 {rules['min']}"
                if "max" in rules and num_val > rules["max"]:
                    return False, f"'{p_meta['label']}' 不能大于 {rules['max']}"
            except ValueError:
                return False, f"'{p_meta['label']}' 必须是有效的数字"

        # 5. 正则表达式检查
        if "regex" in rules:
            if not re.match(rules["regex"], str(value)):
                return False, f"'{p_meta['label']}' 格式不正确"

        # 6. 非空字符串检查 (明确要求非空的内容)
        if rules.get("not_empty") and not str(value).strip():
            return False, f"'{p_meta['label']}' 不能为空"

    return True, ""
