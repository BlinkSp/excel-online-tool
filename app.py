import streamlit as st
import pandas as pd
import re
import zipfile

from io import BytesIO


# ============================================================
# 页面设置
# ============================================================

st.set_page_config(
    page_title="Excel 在线工具箱",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel 在线工具箱")

st.write(
    "支持 Excel 批量合并、数据清洗、去重、按字段拆分等功能。"
)

st.caption("当前版本：V0.3")


# ============================================================
# 工具函数
# ============================================================

def get_sheet_names(file):
    """获取 Excel 中的所有 Sheet"""

    try:
        file.seek(0)

        excel_file = pd.ExcelFile(file)

        sheets = excel_file.sheet_names

        excel_file.close()

        file.seek(0)

        return sheets

    except Exception:
        return []


def load_excel_files(
    files,
    sheet_mode,
    selected_sheet=None,
    add_source=True
):
    """
    批量读取 Excel
    """

    all_data = []

    errors = []

    structures = []

    for file in files:

        try:

            file.seek(0)

            excel_file = pd.ExcelFile(file)

            sheets = excel_file.sheet_names

            # -----------------------------------------
            # 选择需要读取的 Sheet
            # -----------------------------------------

            if sheet_mode == "每个文件第一个工作表":

                target_sheets = [sheets[0]]

            elif sheet_mode == "指定同名工作表":

                if selected_sheet not in sheets:

                    errors.append(
                        f"{file.name}：不存在工作表 "
                        f"「{selected_sheet}」"
                    )

                    continue

                target_sheets = [selected_sheet]

            else:

                # 读取所有工作表
                target_sheets = sheets


            # -----------------------------------------
            # 读取
            # -----------------------------------------

            for sheet in target_sheets:

                df = excel_file.parse(
                    sheet_name=sheet
                )

                original_columns = list(
                    map(str, df.columns)
                )

                structures.append(
                    {
                        "文件": file.name,
                        "Sheet": sheet,
                        "字段": original_columns
                    }
                )

                if add_source:

                    df["来源文件"] = file.name

                    df["来源Sheet"] = sheet

                all_data.append(df)

            excel_file.close()

        except Exception as e:

            errors.append(
                f"{file.name}：{str(e)}"
            )


    if all_data:

        result = pd.concat(
            all_data,
            ignore_index=True,
            sort=False
        )

    else:

        result = pd.DataFrame()

    return result, errors, structures


def safe_filename(value):
    """
    把字段值转换成合法文件名
    """

    if pd.isna(value):
        value = "空值"

    value = str(value).strip()

    if not value:
        value = "空值"

    value = re.sub(
        r'[\\/:*?"<>|]',
        "_",
        value
    )

    return value[:80]


def safe_sheet_name(value):
    """
    Excel Sheet 名称最多31个字符
    """

    name = safe_filename(value)

    return name[:31]


# ============================================================
# 上传文件
# ============================================================

st.subheader("① 上传 Excel")

uploaded_files = st.file_uploader(
    "请选择一个或多个 Excel 文件",
    type=["xlsx"],
    accept_multiple_files=True
)


if uploaded_files:

    st.success(
        f"已上传 {len(uploaded_files)} 个 Excel 文件"
    )

    with st.expander("查看已上传文件"):

        for file in uploaded_files:

            sheets = get_sheet_names(file)

            st.write(
                f"📄 {file.name} ｜ Sheet：{', '.join(sheets)}"
            )


    # ========================================================
    # Sheet 设置
    # ========================================================

    st.subheader("② Sheet 设置")

    sheet_mode = st.radio(
        "请选择读取方式",
        [
            "每个文件第一个工作表",
            "指定同名工作表",
            "读取所有工作表"
        ],
        horizontal=True
    )


    selected_sheet = None


    if sheet_mode == "指定同名工作表":

        all_sheet_sets = []

        for file in uploaded_files:

            sheets = get_sheet_names(file)

            if sheets:
                all_sheet_sets.append(
                    set(sheets)
                )


        if all_sheet_sets:

            common_sheets = sorted(
                set.intersection(
                    *all_sheet_sets
                )
            )

        else:

            common_sheets = []


        if common_sheets:

            selected_sheet = st.selectbox(
                "请选择所有文件共有的 Sheet",
                common_sheets
            )

        else:

            st.error(
                "这些 Excel 没有共同的 Sheet 名称。"
            )


    add_source = st.checkbox(
        "增加“来源文件”和“来源Sheet”字段",
        value=True
    )


    # ========================================================
    # 预读取
    # ========================================================

    can_load = not (
        sheet_mode == "指定同名工作表"
        and selected_sheet is None
    )


    if can_load:

        raw_data, errors, structures = load_excel_files(
            uploaded_files,
            sheet_mode,
            selected_sheet,
            add_source
        )


        # ====================================================
        # 异常提示
        # ====================================================

        if errors:

            st.warning("发现文件异常：")

            for error in errors:
                st.write("⚠️", error)


        if not raw_data.empty:

            # 判断字段结构是否一致

            if structures:

                first_columns = structures[0]["字段"]

                different_structure = []

                for item in structures[1:]:

                    if item["字段"] != first_columns:

                        different_structure.append(item)


                if different_structure:

                    st.warning(
                        "检测到部分 Excel 字段结构不完全一致。"
                        "系统仍会合并，不存在的字段会自动留空。"
                    )


            # =================================================
            # 功能 Tabs
            # =================================================

            tab1, tab2 = st.tabs(
                [
                    "🔗 Excel 合并与清洗",
                    "✂️ Excel 拆分"
                ]
            )


            # =================================================
            # TAB 1：合并 + 清洗
            # =================================================

            with tab1:

                st.subheader(
                    "③ 合并与数据清洗"
                )

                st.write(
                    f"当前读取到 **{len(raw_data)}** 条原始记录。"
                )


                remove_empty_rows = st.checkbox(
                    "删除完全空白的数据行",
                    value=True,
                    key="remove_empty_merge"
                )


                dedupe_mode = st.selectbox(
                    "去重方式",
                    [
                        "不去重",
                        "整行完全重复",
                        "按指定字段去重"
                    ]
                )


                dedupe_columns = []


                metadata_columns = [
                    "来源文件",
                    "来源Sheet"
                ]

                selectable_columns = [
                    col
                    for col in raw_data.columns
                    if col not in metadata_columns
                ]


                if dedupe_mode == "按指定字段去重":

                    dedupe_columns = st.multiselect(
                        "请选择用于判断重复的字段",
                        selectable_columns
                    )


                keep_mode_text = st.selectbox(
                    "重复数据保留方式",
                    [
                        "保留第一条",
                        "保留最后一条"
                    ]
                )

                keep_mode = (
                    "first"
                    if keep_mode_text == "保留第一条"
                    else "last"
                )


                if st.button(
                    "🚀 开始合并和清洗",
                    type="primary"
                ):

                    result = raw_data.copy()

                    original_count = len(result)


                    # -----------------------------------------
                    # 删除空行
                    # -----------------------------------------

                    if remove_empty_rows:

                        data_columns = [
                            col
                            for col in result.columns
                            if col not in metadata_columns
                        ]

                        result = result.dropna(
                            how="all",
                            subset=data_columns
                        )


                    after_empty_count = len(result)


                    # -----------------------------------------
                    # 去重
                    # -----------------------------------------

                    if dedupe_mode == "整行完全重复":

                        compare_columns = [
                            col
                            for col in result.columns
                            if col not in metadata_columns
                        ]

                        result = result.drop_duplicates(
                            subset=compare_columns,
                            keep=keep_mode
                        )


                    elif (
                        dedupe_mode
                        == "按指定字段去重"
                    ):

                        if dedupe_columns:

                            result = result.drop_duplicates(
                                subset=dedupe_columns,
                                keep=keep_mode
                            )

                        else:

                            st.warning(
                                "你选择了“按指定字段去重”，"
                                "但还没有选择字段。"
                            )


                    result = result.reset_index(
                        drop=True
                    )


                    st.session_state[
                        "clean_result"
                    ] = result


                    # -----------------------------------------
                    # 统计
                    # -----------------------------------------

                    removed_empty = (
                        original_count
                        - after_empty_count
                    )

                    removed_duplicates = (
                        after_empty_count
                        - len(result)
                    )


                    st.success(
                        "数据处理完成！"
                    )

                    col1, col2, col3 = st.columns(3)

                    col1.metric(
                        "原始数据",
                        original_count
                    )

                    col2.metric(
                        "删除空行",
                        removed_empty
                    )

                    col3.metric(
                        "删除重复数据",
                        removed_duplicates
                    )


                # ---------------------------------------------
                # 展示合并结果
                # ---------------------------------------------

                if (
                    "clean_result"
                    in st.session_state
                ):

                    result = st.session_state[
                        "clean_result"
                    ]

                    st.subheader(
                        "④ 处理结果预览"
                    )

                    col1, col2 = st.columns(2)

                    col1.metric(
                        "最终数据行数",
                        len(result)
                    )

                    col2.metric(
                        "字段数量",
                        len(result.columns)
                    )

                    st.dataframe(
                        result,
                        use_container_width=True,
                        height=500
                    )


                    # -----------------------------------------
                    # Excel 导出
                    # -----------------------------------------

                    output = BytesIO()

                    with pd.ExcelWriter(
                        output,
                        engine="openpyxl"
                    ) as writer:

                        result.to_excel(
                            writer,
                            sheet_name="处理结果",
                            index=False
                        )

                    output.seek(0)


                    st.download_button(
                        label="📥 下载处理后的 Excel",
                        data=output,
                        file_name="Excel处理结果.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        )
                    )


            # =================================================
            # TAB 2：拆分 Excel
            # =================================================

            with tab2:

                st.subheader(
                    "③ 按字段拆分 Excel"
                )

                st.write(
                    "可以按照负责人、地区、项目状态、"
                    "部门等任意字段自动拆分。"
                )


                split_source = raw_data.copy()


                split_columns = [
                    col
                    for col in split_source.columns
                    if col not in metadata_columns
                ]


                if split_columns:

                    split_column = st.selectbox(
                        "请选择拆分依据",
                        split_columns
                    )


                    remove_empty_split = st.checkbox(
                        "拆分前删除完全空白行",
                        value=True,
                        key="remove_empty_split"
                    )


                    export_mode = st.radio(
                        "拆分结果保存方式",
                        [
                            "多个 Excel 文件打包为 ZIP",
                            "一个 Excel 内生成多个 Sheet"
                        ]
                    )


                    if st.button(
                        "✂️ 开始拆分",
                        type="primary"
                    ):

                        split_result = (
                            split_source.copy()
                        )


                        if remove_empty_split:

                            data_columns = [
                                col
                                for col in split_result.columns
                                if col not in metadata_columns
                            ]

                            split_result = (
                                split_result.dropna(
                                    how="all",
                                    subset=data_columns
                                )
                            )


                        # 空值也保留为一组

                        split_result[
                            "__拆分字段__"
                        ] = (
                            split_result[
                                split_column
                            ]
                            .fillna("空值")
                            .astype(str)
                            .replace("", "空值")
                        )


                        groups = list(
                            split_result.groupby(
                                "__拆分字段__",
                                dropna=False
                            )
                        )


                        st.success(
                            f"共拆分为 {len(groups)} 组。"
                        )


                        preview = pd.DataFrame(
                            {
                                split_column: [
                                    group_name
                                    for group_name, _
                                    in groups
                                ],

                                "数据数量": [
                                    len(group_df)
                                    for _, group_df
                                    in groups
                                ]
                            }
                        )

                        st.dataframe(
                            preview,
                            use_container_width=True
                        )


                        # =====================================
                        # 模式一：ZIP
                        # =====================================

                        if (
                            export_mode
                            == "多个 Excel 文件打包为 ZIP"
                        ):

                            zip_buffer = BytesIO()

                            used_names = {}


                            with zipfile.ZipFile(
                                zip_buffer,
                                mode="w",
                                compression=zipfile.ZIP_DEFLATED
                            ) as zip_file:

                                for (
                                    group_name,
                                    group_df
                                ) in groups:

                                    group_df = (
                                        group_df
                                        .drop(
                                            columns=[
                                                "__拆分字段__"
                                            ]
                                        )
                                    )


                                    filename = (
                                        safe_filename(
                                            group_name
                                        )
                                    )


                                    # 防止文件名重复

                                    if filename in used_names:

                                        used_names[
                                            filename
                                        ] += 1

                                        filename = (
                                            f"{filename}_"
                                            f"{used_names[filename]}"
                                        )

                                    else:

                                        used_names[
                                            filename
                                        ] = 1


                                    excel_buffer = BytesIO()


                                    with pd.ExcelWriter(
                                        excel_buffer,
                                        engine="openpyxl"
                                    ) as writer:

                                        group_df.to_excel(
                                            writer,
                                            sheet_name="数据",
                                            index=False
                                        )


                                    excel_buffer.seek(0)


                                    zip_file.writestr(
                                        f"{filename}.xlsx",
                                        excel_buffer.getvalue()
                                    )


                            zip_buffer.seek(0)


                            st.download_button(
                                label="📦 下载拆分结果 ZIP",
                                data=zip_buffer,
                                file_name=(
                                    f"按{split_column}拆分结果.zip"
                                ),
                                mime="application/zip"
                            )


                        # =====================================
                        # 模式二：多个Sheet
                        # =====================================

                        else:

                            excel_output = BytesIO()

                            used_sheet_names = {}


                            with pd.ExcelWriter(
                                excel_output,
                                engine="openpyxl"
                            ) as writer:

                                for (
                                    group_name,
                                    group_df
                                ) in groups:

                                    group_df = (
                                        group_df
                                        .drop(
                                            columns=[
                                                "__拆分字段__"
                                            ]
                                        )
                                    )


                                    sheet_name = (
                                        safe_sheet_name(
                                            group_name
                                        )
                                    )


                                    # 防止 Sheet 重名

                                    original_name = (
                                        sheet_name
                                    )

                                    counter = 1

                                    while (
                                        sheet_name
                                        in used_sheet_names
                                    ):

                                        counter += 1

                                        suffix = (
                                            f"_{counter}"
                                        )

                                        sheet_name = (
                                            original_name[
                                                :31-len(suffix)
                                            ]
                                            + suffix
                                        )


                                    used_sheet_names[
                                        sheet_name
                                    ] = True


                                    group_df.to_excel(
                                        writer,
                                        sheet_name=sheet_name,
                                        index=False
                                    )


                            excel_output.seek(0)


                            st.download_button(
                                label="📥 下载拆分后的 Excel",
                                data=excel_output,
                                file_name=(
                                    f"按{split_column}拆分结果.xlsx"
                                ),
                                mime=(
                                    "application/vnd."
                                    "openxmlformats-officedocument."
                                    "spreadsheetml.sheet"
                                )
                            )


                else:

                    st.warning(
                        "没有可以用于拆分的数据字段。"
                    )


        else:

            st.error(
                "没有成功读取到 Excel 数据。"
            )


else:

    st.info(
        "请先上传 Excel 文件。"
    )