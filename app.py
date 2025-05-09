import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import os
import xlrd
import matplotlib.font_manager as fm

def load_excel_file(file_path):
    """
    兼容 xlsx 和 xls 文件读取
    """
    if file_path.endswith('.xls'):
        # 旧格式 xls 文件使用 xlrd 进行读取
        try:
            return pd.ExcelFile(file_path, engine='xlrd')
        except Exception as e:
            st.error(f"读取 .xls 文件出错: {e}")
            return None
    else:
        # xlsx 文件使用 openpyxl 进行读取
        try:
            return pd.ExcelFile(file_path)
        except Exception as e:
            st.error(f"读取 .xlsx 文件出错: {e}")
            return None

def convert_date(date_series):
    """尝试多种日期格式，将日期统一转换为标准格式"""
    formats = ["%Y%m%d", "%Y-%m-%d", "%Y/%m/%d"]
    
    for fmt in formats:
        try:
            # 直接尝试转换，如果成功则返回转换后的结果
            converted = pd.to_datetime(date_series, format=fmt, errors='coerce')
            if converted.notna().sum() > 0:
                return converted
        except:
            continue
    
    # 如果所有格式均未成功，则进行自动推断
    return pd.to_datetime(date_series, errors='coerce')

#设置全局属性
config ={
"font.family":'Simhei',
}
#更新全局属性配置
plt.rcParams.update(config)

# 防止负号显示为方块
matplotlib.rcParams['axes.unicode_minus'] = False

# matplotlib.rcParams['font.family'] = 'SimHei'
# matplotlib.rcParams['axes.unicode_minus'] = False

# # 添加字体路径
# font_path = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
# fm.fontManager.addfont(font_path)

# # 重新构建字体缓存
# fm._rebuild()

# plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei']
# plt.rcParams['axes.unicode_minus'] = False

st.set_page_config(page_title="悦海盈和产品可视化", layout="wide")
st.title("📈 悦海盈和产品可视化平台")

# 分类标签结构
categories = {
    "中性策略": ["悦海盈和2号", "悦海盈和7号", "悦海盈和18号"],
    "混合对冲": ["悦海盈和春晓1号"],
    "精选对冲": ["悦海盈和精选对冲2号1期"],
    "灵活对冲": ["悦海盈和12号"],
    "量化选股": ["悦海盈和量化选股1号"],
    "指数增强": ["悦海盈和500指增1号", "悦海盈和16号", "悦海盈和A500指数增强1号"]
}

st.markdown("通过策略分类浏览多个产品的净值和业绩表现")

# 管理模式控制上传入口
admin_mode = st.query_params.get("admin", "false").lower() == "true"
if admin_mode:
    st.sidebar.title("🛠 数据上传（仅管理员）")
    uploaded_product = st.sidebar.selectbox("选择产品名称", sum(categories.values(), []))
    upload = st.sidebar.file_uploader("上传 Excel 文件", type=["xlsx","xls"])
    if upload:
        save_name = {
            "悦海盈和2号": "2号.xlsx",
            "悦海盈和7号": "7号.xlsx",
            "悦海盈和18号": "18号.xlsx",
            "悦海盈和春晓1号": "春晓1号.xlsx",
            "悦海盈和精选对冲2号1期": "精选对冲2号1期.xlsx",
            "悦海盈和12号": "12号.xlsx",
            "悦海盈和量化选股1号": "量化选股1号.xlsx",
            "悦海盈和500指增1号": "500指增回测业绩.xlsx",
            "悦海盈和16号": "悦海16号-1000指增业绩.xlsx",
            "悦海盈和A500指数增强1号": "A500指增1号.xlsx"
        }.get(uploaded_product, "product.xlsx")
        with open(save_name, "wb") as f:
            f.write(upload.read())
        st.sidebar.success(f"{uploaded_product} 数据已更新 ✅")
selected_category = st.selectbox("📂 请选择策略分类：", list(categories.keys()))
if categories[selected_category]:
    selected_product = st.selectbox("📌 请选择产品：", categories[selected_category])
    code_map = {
        "悦海盈和2号": "2号.xlsx",
        "悦海盈和7号": "7号.xlsx",
        "悦海盈和18号": "18号.xlsx",
        "悦海盈和春晓1号": "春晓1号.xlsx",
        "悦海盈和精选对冲2号1期": "精选对冲2号1期.xlsx",
        "悦海盈和12号": "12号.xlsx",
        "悦海盈和量化选股1号": "量化选股1号.xlsx",
        "悦海盈和500指增1号":"500指增回测业绩.xlsx",
        "悦海盈和16号":"悦海16号-1000指增业绩.xlsx",
        "悦海盈和A500指数增强1号": "A500指增1号.xlsx"
    }
    start_date_map = {
        "悦海盈和2号": "2021年3月5日",
        "悦海盈和7号": "2021年9月7日",
        "悦海盈和18号": "2022年11月15日",
        "悦海盈和春晓1号": "2023年4月28日",
        "悦海盈和精选对冲2号1期": "2025年2月26日",
        "悦海盈和12号": "2022年4月28日",
        "悦海盈和量化选股1号": "2025年1月24日",
        "悦海盈和500指增1号":"2022年2月28日",
        "悦海盈和16号":"2023年4月27日",
        "悦海盈和A500指数增强1号": "2024年12月27日"
    }
    file_name = code_map.get(selected_product)

    if file_name and os.path.exists(file_name):
        #excel = pd.ExcelFile(file_name)
        excel = load_excel_file(file_name)
        if selected_product in ["悦海盈和精选对冲2号1期", "悦海盈和A500指数增强1号"]:
            df_weekly = excel.parse("周度")
            df_plot = excel.parse("周报图")
            df_raw_monthly = excel.parse("月度", header=None)
            df_plot.columns = df_plot.columns.str.strip()
            df_plot["日期"] = pd.to_datetime(df_plot["净值日期"])
            df_plot = df_plot.sort_values("日期")
        elif selected_product in ["悦海盈和500指增1号", "悦海盈和16号"]:
            df_weekly = excel.parse("拼接周度")
            if selected_product == "悦海盈和500指增1号" and "日期" not in df_weekly.columns:
                df_weekly["日期"] = pd.to_datetime(df_weekly.iloc[1:, 0], errors='coerce').reset_index(drop=True)                     
            df_plot = excel.parse("周报图")
            df_plot.columns = df_plot.columns.str.strip()
            if "日期" not in df_plot.columns:
                df_plot.insert(0, "日期", convert_date(df_plot.iloc[1:, 0]))          
            else:
                df_plot["日期"] = convert_date(df_plot["日期"])   
            df_raw_monthly = excel.parse("拼接月度", header=None)
            df_plot.columns = df_plot.columns.str.strip()
            
            df_plot["日期"] = convert_date(df_plot["日期"])
            df_plot = df_plot.dropna(subset=["日期"]).sort_values("日期").reset_index(drop=True)
            
            # 找到第一个 "累计超额" 列，避免使用 .1 或 .2
            excess_cols = [col for col in df_plot.columns if "累计超额" in col]
            if len(excess_cols) > 0:
                excess_col = excess_cols[0]  # 选择第一个 "累计超额" 列
            else:
                st.warning("⚠️ '累计超额' 列未找到，无法绘制灰色区域。")
                excess_col = None    
        else:
            has_daily = "日度" in excel.sheet_names
            df_weekly = excel.parse("周度")
            df_raw_monthly = excel.parse("月度", header=None)
            df_daily = excel.parse("日度") if has_daily else None

        df_weekly.columns = df_weekly.columns.str.strip()
        date_col = "净值日期" if selected_product in ["悦海盈和春晓1号", "悦海盈和精选对冲2号1期", "悦海盈和12号", "悦海盈和量化选股1号","悦海盈和A500指数增强1号"] else "日期"
        df_weekly["日期"] = pd.to_datetime(df_weekly[date_col])
        df_weekly = df_weekly.sort_values("日期")
#         report_date = df_weekly['日期'].max().strftime("%Y-%m-%d")
#         st.markdown(f"📅 报告截至日期：**{report_date}**")
        # 报告截至日期逻辑修改
        if "日度" in excel.sheet_names:
            try:
                df_daily = excel.parse("日度")
                df_daily["日期"] = pd.to_datetime(df_daily.iloc[:, 0], errors='coerce')
                df_daily = df_daily.dropna(subset=["日期"])
                latest_daily_date = df_daily["日期"].max()
            except Exception as e:
                st.warning(f"读取日度数据出错：{e}")
                latest_daily_date = pd.NaT
        else:
            latest_daily_date = pd.NaT

        # 获取周度数据的最新日期
        latest_weekly_date = df_weekly["日期"].max()

        # 选择较晚的日期作为报告截至日期
        if pd.notna(latest_daily_date) and pd.notna(latest_weekly_date):
            report_date = max(latest_daily_date, latest_weekly_date).strftime("%Y-%m-%d")
        elif pd.notna(latest_daily_date):
            report_date = latest_daily_date.strftime("%Y-%m-%d")
        else:
            report_date = latest_weekly_date.strftime("%Y-%m-%d")

        st.markdown(f"📅 报告截至日期：**{report_date}**")

        if selected_product == "悦海盈和2号":
            rate_col = "增长率"
            index_cols = ["中证500指数"]
            product_col = "悦海盈和2号"
        elif selected_product == "悦海盈和7号":
            rate_col = "周增长率"
            index_cols = ["中证500指数"]
            product_col = "悦海盈和7号"
        elif selected_product == "悦海盈和18号":
            rate_col = "周增长率"
            index_cols = ["中证500指数", "中证1000表现"]
            product_col = "悦海盈和18号" 
        elif selected_product == "悦海盈和春晓1号":
            rate_col = "周增长率"
            index_cols = ["中证500表现", "中证1000表现"]
            product_col = "春晓1号"
        elif selected_product == "悦海盈和精选对冲2号1期":
            rate_col = "周增长率"
            index_cols = ["中证500", "中证1000"]
            product_col_plot = "精选对冲2号1期"
            product_col_nav = "精选对冲2号"
        elif selected_product == "悦海盈和12号":
            rate_col = "周增长率"
            index_cols = ["中证500指数", "中证1000指数"]
            product_col = "悦海盈和12号"
        elif selected_product == "悦海盈和量化选股1号":
            rate_col = "周增长率"
            index_cols = ["中证全指.1"]
            product_col = "量化选股1号"
        elif selected_product == "悦海盈和500指增1号":
            rate_col = "拼接业绩周增长率"
            product_col = "500指数增强1号"
            index_cols = ["500指数收益"]
            backtest_col = "指数增强回测业绩"
            excess_col = "累计超额"
        elif selected_product == "悦海盈和16号":
            rate_col = "拼接指增收益率"
            product_col = "悦海盈和16号"
            index_cols = ["1000指数收益"]
            backtest_col = "指增回测业绩"
            excess_col = "累计超额"
        elif selected_product == "悦海盈和A500指数增强1号":
            rate_col = "增长率"
            product_col = "指增产品净值"
            index_cols = ["A500指数"]
            excess_col = "累计超额"
        else:
            rate_col = "周增长率"
            index_cols = ["中证500指数"]
            product_col = selected_product

        df_weekly[rate_col] = pd.to_numeric(df_weekly[rate_col], errors='coerce')
        df_weekly['回撤'] = pd.to_numeric(df_weekly['回撤'], errors='coerce')

        if selected_product == "悦海盈和精选对冲2号1期":
            df_weekly[product_col_nav] = pd.to_numeric(df_weekly[product_col_nav], errors='coerce')
            latest_nav = df_weekly[product_col_nav].iloc[-1]
            first_nav = df_weekly[product_col_nav].iloc[0]
        else:
            df_weekly[product_col] = pd.to_numeric(df_weekly[product_col], errors='coerce')
            latest_nav = df_weekly[product_col].iloc[-1]
            first_nav = df_weekly[product_col].iloc[0]

        last_week_ret = df_weekly[rate_col].iloc[-1]
        cumulative_ret = latest_nav / first_nav - 1 if first_nav != 0 else np.nan
        max_dd = -df_weekly['回撤'].min()
        ann_vol = df_weekly[rate_col].std() * np.sqrt(52)
        avg_weekly_ret = df_weekly[rate_col].mean()
        ann_return = (1 + avg_weekly_ret) ** 52 - 1
        sharpe = ann_return / ann_vol if ann_vol != 0 else np.nan
        
        if selected_product == "悦海盈和精选对冲2号1期":
            
            # 固定信息数据
            fixed_data = [
                ["产品编码", "SASQ94", "份额锁定", "180天"],
                ["产品策略", "精选对冲", "投资经理", "张晴"],
                ["产品结构", "母子结构", "基金状态", "正在运作"],
                ["管理人", "青岛悦海盈和基金投资管理有限公司", "最新净值", f"{latest_nav:.4f}"],
                ["托管机构", "国泰君安证券股份有限公司", "本周收益", f"{last_week_ret * 100:.2f}%"],
                ["成立日", "2025/2/25", "最大回撤", f"{max_dd * 100:.2f}%"],
                ["建仓日", "2025/3/26", "累计收益", f"{cumulative_ret * 100:.2f}%"]
            ]
            
            st.markdown("""
            **策略描述—精选对冲策略**  
            ✓ 对标指数：中证500指数、中证1000指数。  
            ✓ 空头端：股指期货。  
            ✓ 策略逻辑：基于对标指数构建择时模型，预测1-3天收益，根据预测结果决定敞口，敞口暴露范围为0-100%。非满仓状态下会根据基差状态决定是否对冲，基差处于极端情况时选择降低仓位。
            """)

            # 列标题直接作为第一行数据，不使用 `st.table` 的 `.style.hide(axis="index")`
            columns = ["", "", "", ""]
            fixed_df = pd.DataFrame(fixed_data, columns=columns)

            # 将 DataFrame 转换为 HTML 表格，隐藏索引
            html_table = fixed_df.to_html(index=False, header=False, border=0)

            # 显示表格
#             st.subheader("📋 基金信息")
#             st.markdown(html_table, unsafe_allow_html=True)

            # 动态数据改为 `dict` 类型
            perf_data = {
                "最新净值": f"{latest_nav:.4f}",
                "本周收益": f"{last_week_ret * 100:.2f}%",
                "最大回撤": f"{max_dd * 100:.2f}%",
                "累计收益": f"{cumulative_ret * 100:.2f}%"
            }         
         
        elif selected_product == "悦海盈和量化选股1号":
            perf_data = {
                "运作起始日": start_date_map[selected_product],
                "最新净值": f"{latest_nav:.4f}",
                "本周收益": f"{last_week_ret * 100:.2f}%",
                "最大回撤": f"{max_dd * 100:.2f}%",
                "累计收益": f"{cumulative_ret * 100:.2f}%"
            }
            perf_df = pd.DataFrame(perf_data.items(), columns=["指标", selected_product])

            st.markdown("""
            **策略描述—量化选股策略**  
            ✓ 全市场选股：全市场轮动选股，不严格对标某个指数，发挥模型选股、创造超额收益的原始能力。  
            ✓ 仓位管理：构建择时模型，根据模型的多空信号动态调整仓位。模型给出空头信号时降低仓位，降低产品波动。
            """)
            
        elif selected_product == "悦海盈和500指增1号":
            df_weekly[rate_col] = pd.to_numeric(df_weekly[rate_col], errors='coerce')
            last_week_return = df_weekly[rate_col].iloc[-1]
            last_excess = df_weekly["周超额"].iloc[-1]
            avg_excess = df_weekly["周超额"].mean()
            ann_excess = (1 + avg_excess)**52 - 1
            ann_excess_vol = df_weekly["周超额"].std() * np.sqrt(52)
            ann_return = (1 + df_weekly[rate_col].mean())**52 - 1
            max_dd = -df_weekly["回撤"].min()
            sharpe = ann_return / (df_weekly[rate_col].std() * np.sqrt(52))
            
            perf_data = {
            "运作起始日": start_date_map[selected_product],
            "最新净值": f"{latest_nav:.4f}",
            "本周收益": f"{last_week_return * 100:.2f}%",
            "本周超额": f"{last_excess * 100:.2f}%",
            "年化超额收益": f"{ann_excess * 100:.2f}%",
            "年化超额波动率": f"{ann_excess_vol * 100:.2f}%",
            "年化收益": f"{ann_return * 100:.2f}%",
            "最大回撤": f"{max_dd * 100:.2f}%",
            "夏普比率": f"{sharpe:.2f}"
        }
       
        elif selected_product == "悦海盈和16号":
            df_weekly[rate_col] = pd.to_numeric(df_weekly[rate_col], errors='coerce')
            last_week_return = df_weekly[rate_col].iloc[-1]
            last_excess = df_weekly["周超额"].iloc[-1]
            avg_excess = df_weekly["周超额"].mean()
            ann_excess = (1 + avg_excess)**52 - 1
            ann_excess_vol = df_weekly["周超额"].std() * np.sqrt(52)
            ann_return = (1 + df_weekly[rate_col].mean())**52 - 1
            max_dd = -df_weekly["回撤"].min()
            excess_sharpe = ann_excess / ann_excess_vol if ann_excess_vol != 0 else np.nan
            
            perf_data = {
            "运作起始日": start_date_map[selected_product],
            "最新净值": f"{latest_nav:.4f}",
            "本周收益": f"{last_week_return * 100:.2f}%",
            "本周超额": f"{last_excess * 100:.2f}%",
            "年化超额收益": f"{ann_excess * 100:.2f}%",
            "年化超额波动率": f"{ann_excess_vol * 100:.2f}%",
            "年化收益": f"{ann_return * 100:.2f}%",
            "最大回撤": f"{max_dd * 100:.2f}%",
            "超额夏普比率": f"{excess_sharpe:.2f}"
        }
        elif selected_product == "悦海盈和A500指数增强1号":
            df_weekly[rate_col] = pd.to_numeric(df_weekly[rate_col], errors='coerce')
            last_week_return = df_weekly[rate_col].iloc[-1]
            last_excess = df_weekly["周度超额"].iloc[-1]
            latest_nav = df_weekly[product_col].iloc[-1]
            first_nav = df_weekly[product_col].iloc[0]
            cumulative_ret = latest_nav / first_nav - 1 if first_nav != 0 else np.nan

            # 计算运作以来指数收益
            last_index = df_weekly["A500指数"].iloc[-1]
            first_index = df_weekly["A500指数"].iloc[0]
            index_ret = last_index / first_index - 1 if first_index != 0 else np.nan

            # 运作以来超额收益
            cumulative_excess_ret = cumulative_ret - index_ret

            # 计算超额最大回撤
            max_excess_dd = -df_weekly["周度超额"].min()

            perf_data = {
                "运作起始日": start_date_map[selected_product],
                "最新净值": f"{latest_nav:.4f}",
                "本周收益": f"{last_week_return * 100:.2f}%",
                "本周超额收益": f"{last_excess * 100:.2f}%",
                "运作以来收益": f"{cumulative_ret * 100:.2f}%",
                "运作以来超额收益": f"{cumulative_excess_ret * 100:.2f}%",
                "超额最大回撤": f"{max_excess_dd * 100:.2f}%"
            }

        else:
            perf_data = {
                "运作起始日": start_date_map[selected_product],
                "最新净值": f"{latest_nav:.4f}",
                "本周收益": f"{last_week_ret * 100:.2f}%",
                "年化收益": f"{ann_return * 100:.2f}%",
                "年化波动率": f"{ann_vol * 100:.2f}%",
                "最大回撤": f"{max_dd * 100:.2f}%",
                "年化夏普比率": f"{sharpe:.2f}",
                "累计收益": f"{cumulative_ret * 100:.2f}%"
            }

        perf_df = pd.DataFrame(perf_data.items(), columns=["指标", selected_product])

        if selected_product == "悦海盈和春晓1号":
            st.markdown("""
            **策略描述—混合对冲策略**  
            ✓ 对标指数：中证500指数、中证1000指数。  
            ✓ 空头端：股指期货。  
            ✓ 策略逻辑：80%中性策略+20%灵活对冲策略。多头端持有一揽子股票，建立分散化投资组合，通过量化模型获取超额收益alpha。空头端持有对冲工具，灵活对冲部分基于对标指数构建择时模型，预测1-3天收益，根据预测结果决定敞口。
            """)
        elif selected_product == "悦海盈和12号":
            st.markdown("""
            **策略描述—灵活对冲策略**  
            ✓ 对标指数：中证500指数、中证1000指数。  
            ✓ 空头端：股指期货。  
            ✓ 策略逻辑：基于对标指数构建择时模型，预测1-3天收益，根据预测结果决定敞口，敞口暴露范围为0-100%。非满仓状态下会根据基差状态决定是否对冲，基差处于极端情况时选择降低仓位。
            """)

        col1, col2 = st.columns(2)

        with col1:
            if selected_product == "悦海盈和精选对冲2号1期":
                st.subheader("📋 基金信息")
                st.markdown(html_table, unsafe_allow_html=True)
            else:
                st.subheader("📋 基金业绩表现")
                st.table(perf_df.set_index("指标"))

        with col2:
            st.subheader("📊 净值变化曲线")
            fig, ax = plt.subplots(figsize=(6, 3.5))
            if selected_product == "悦海盈和精选对冲2号1期":
                ax.plot(df_plot['日期'], df_plot[product_col_plot], label=product_col_nav, linewidth=2)
                for idx in index_cols:
                    if idx in df_plot.columns:
                        ax.plot(df_plot['日期'], pd.to_numeric(df_plot[idx], errors='coerce'), label=idx, linestyle='--')
            

            elif selected_product in ["悦海盈和500指增1号", "悦海盈和16号"]:
                df_plot["日期"] = pd.to_datetime(df_plot.iloc[1:, 0], errors='coerce') 
                
                fig, ax1 = plt.subplots(figsize=(7, 4))
                ax1.plot(df_plot['日期'], df_plot[backtest_col], linestyle='--', label="指增回测业绩", color='orange')

                ax1.plot(df_plot['日期'], df_plot[product_col], color='orange', label="500指数增强1号", linewidth=2)
                ax2 = ax1.twinx()
                
                for idx in index_cols:
                    if idx in df_plot.columns:
                        ax1.plot(df_plot['日期'], df_plot[idx], label=idx)
                
                # 坐标轴范围设置
                if selected_product == "悦海盈和500指增1号":
                    ax1.set_ylim(0.6, 2.2)
                    ax1.set_yticks(np.arange(0.6, 2.3, 0.2))
                    ax2.ylim=(0, 0.9)
                    ax2.yticks = np.arange(0, 1.0, 0.1)
                elif selected_product == "悦海盈和16号":
                    ax1.set_ylim(0.4, 1.6)
                    ax1.set_yticks(np.arange(0.4, 1.7, 0.2))
                    ax2.ylim = (0, 0.7)
                    ax2.yticks = np.arange(0, 0.8, 0.1)
        
                ax1.set_ylabel("单位净值")
                ax1.set_xlabel("日期")
                ax1.legend(loc='upper left')

                ax2.fill_between(df_plot['日期'], df_plot[excess_col], color='grey', alpha=0.3, label="累计超额")
                ax2.legend(loc='upper right')
                ax2.set_ylabel("累计超额")
                
#                 ax2.set_ylim(ax2_ylim)
#                 ax2.set_yticks(ax2_yticks)
                
                
                
                fig.autofmt_xdate(rotation=45)
                fig.tight_layout()
                
            elif selected_product == "悦海盈和A500指数增强1号":
                fig, ax1 = plt.subplots(figsize=(7, 4))
                ax1.plot(df_plot['日期'], df_plot[product_col], color='blue', label="指增产品净值", linewidth=2)
                ax1.plot(df_plot['日期'], df_plot["A500指数"], color='orange', label="A500指数")

                ax1.set_ylabel("单位净值")
                ax1.set_ylim(0.8, 1.2)
                ax1.set_yticks(np.arange(0.8, 1.2, 0.05))
                ax1.set_xlabel("日期")
                ax1.legend(loc='upper left')

                ax2 = ax1.twinx()
                ax2.set_ylim(0, 0.18)
                ax2.set_yticks(np.arange(0, 0.18, 0.02))
                ax2.fill_between(df_plot['日期'], df_plot[excess_col], color='grey', alpha=0.3, label="累计超额")
                ax2.legend(loc='upper right')
                ax2.set_ylabel("累计超额")

                fig.autofmt_xdate(rotation=45)
                fig.tight_layout()
    
            else:
                ax.plot(df_weekly['日期'], df_weekly[product_col], label=product_col, linewidth=2)
                for idx in index_cols:
                    plot_label = "中证全指" if idx == "中证全指.1" else idx
                    if idx in df_weekly.columns:
                        ax.plot(df_weekly['日期'], pd.to_numeric(df_weekly[idx], errors='coerce'), label=plot_label, linestyle='--')
            #ax.set_title("产品 vs 指数净值对比")
            ax.set_title(f"{selected_product}净值曲线")
            ax.set_xlabel("日期")
            ax.set_ylabel("单位净值")
            fig.autofmt_xdate(rotation=45)
            ax.legend()
            st.pyplot(fig)
        st.subheader("📆 基金月度收益表")
        month_keywords = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', 'YTD']
        if selected_product in ["悦海盈和500指增1号", "悦海盈和16号"]:
            try:
                df_raw_monthly = excel.parse("月度超额", header=None)
                start_idx = df_raw_monthly[df_raw_monthly.iloc[:, 0].astype(str).str.contains("月收益法：", na=False)].index
                if not start_idx.empty:
                    # 找到起始点后的有效数据行
                    start_row = start_idx[0] + 1
                    end_row = start_row
                    for row in range(start_row, len(df_raw_monthly)):
                        # 如果该行第二个单元格为空或为 NaN，则结束
                        if pd.isna(df_raw_monthly.iloc[row, 1]):
                            break
                        end_row = row
                    
                    df_monthly = df_raw_monthly.iloc[start_row:end_row+1].copy()
                    df_monthly.columns = df_monthly.iloc[0]
                    df_monthly = df_monthly[1:].reset_index(drop=True)
                    df_monthly = df_monthly.dropna(how="all", axis=1)
                    df_monthly = df_monthly.set_index("%")
                    df_monthly = df_monthly.loc[:, month_keywords]
                    def to_percent(x):
                        if pd.isna(x):
                            return ""
                        try:
                            return f"{float(str(x).replace('%', '')) * 100:.2f}%"
                        except:
                            return ""
                        
                    df_monthly = df_monthly.applymap(to_percent)
                    st.markdown("绝对收益")
                    st.dataframe(df_monthly, use_container_width=True)
                else:
                    st.warning("⚠️ 未能找到 '月度收益' 起始点，请检查Excel格式。")
                # 读取月度超额
                # 查找所有 "月度超额"
                excess_indices = df_raw_monthly[df_raw_monthly.iloc[:, 0].astype(str).str.contains("月度超额", na=False)].index

                if len(excess_indices) > 1:
                    # 如果找到多个 "月度超额"，取 "月收益法：" 下方的第二个 "月度超额"
                    target_index = excess_indices[1]
                elif len(excess_indices) == 1:
                    # 如果只找到一个，直接取这个 "月度超额"
                    target_index = excess_indices[0]
                else:
                    target_index = None

                if target_index is not None:
                    # 月度超额数据直接从 "月度超额" 开始读取
                    df_excess = df_raw_monthly.iloc[target_index:target_index + 7].copy()
                    df_excess.columns = df_excess.iloc[0]
                    df_excess = df_excess[1:].reset_index(drop=True)
                    df_excess = df_excess.dropna(how="all", axis=1)
                    df_excess = df_excess.set_index("%")
                    df_excess = df_excess.loc[:, month_keywords]
                    df_excess = df_excess.applymap(to_percent)

                    st.markdown("超额")
                    st.dataframe(df_excess, use_container_width=True)
                else:
                    st.warning("⚠️ 未能找到 '月度超额' 起始点，请检查Excel格式。")

            except Exception as e:
                st.error(f"读取 500指增1号 月度收益或超额表出错: {e}")
        elif selected_product == "悦海盈和A500指数增强1号":
            try:
                df_raw_monthly = excel.parse("月度", header=None)

                # 月度收益读取
                start_idx = df_raw_monthly[df_raw_monthly.iloc[:, 0].astype(str).str.contains("绝对收益", na=False)].index
                if not start_idx.empty:
                    start_row = start_idx[0] + 1
                    end_row = start_row

                    # 找到数据结束行
                    for row in range(start_row, len(df_raw_monthly)):
                        if pd.isna(df_raw_monthly.iloc[row, 0]):
                            break
                        end_row = row

                    # 获取列名
                    columns = df_raw_monthly.iloc[start_idx[0]].tolist()

                    # 读取数据区域
                    df_monthly = df_raw_monthly.iloc[start_row:end_row + 1].copy()
                    df_monthly.columns = columns

                    # 保持完整列结构并去除全空行
                    df_monthly = df_monthly.dropna(how="all", axis=0)
                    df_monthly = df_monthly.reindex(columns=columns, fill_value="")

                    # 设置索引列
                    df_monthly.set_index("绝对收益", inplace=True)

                    # 格式化数值，乘以100并加上百分号
                    def format_percent(x):
                        try:
                            if pd.isna(x) or x == "":
                                return ""
                            return f"{float(x) * 100:.2f}%"
                        except:
                            return ""

                    df_monthly = df_monthly.applymap(format_percent)

                    st.markdown("绝对收益")
                    st.dataframe(df_monthly, use_container_width=True)

                # 超额收益读取
                start_idx = df_raw_monthly[df_raw_monthly.iloc[:, 0].astype(str).str.contains("超额", na=False)].index
                if not start_idx.empty:
                    start_row = start_idx[0] + 1
                    end_row = start_row

                    # 找到数据结束行
                    for row in range(start_row, len(df_raw_monthly)):
                        if pd.isna(df_raw_monthly.iloc[row, 0]):
                            break
                        end_row = row

                    # 获取列名
                    columns = df_raw_monthly.iloc[start_idx[0]].tolist()

                    # 读取数据区域
                    df_excess = df_raw_monthly.iloc[start_row:end_row + 1].copy()
                    df_excess.columns = columns

                    # 保持完整列结构并去除全空行
                    df_excess = df_excess.dropna(how="all", axis=0)
                    df_excess = df_excess.reindex(columns=columns, fill_value="")

                    # 设置索引列
                    df_excess.set_index("超额", inplace=True)

                    # 格式化数值
                    df_excess = df_excess.applymap(format_percent)

                    st.markdown("超额")
                    st.dataframe(df_excess, use_container_width=True)

            except Exception as e:
                st.error(f"读取 A500指数增强1号 月度收益或超额收益表出错: {e}")

                             
                            
        else:        
            start_idx = df_raw_monthly[df_raw_monthly.apply(lambda row: all(k in row.values for k in month_keywords), axis=1)].index
            if not start_idx.empty:
                df_monthly = df_raw_monthly.iloc[start_idx[0]:].copy()
                df_monthly.columns = df_monthly.iloc[0]
                df_monthly = df_monthly[1:]
                df_monthly = df_monthly.set_index(df_monthly.columns[0])

                def to_percent(x):
                    if pd.isna(x):
                        return ""
                    try:
                        return f"{float(str(x).replace('%', '')) * 100:.2f}%"
                    except:
                        return ""

                df_monthly_display = df_monthly.applymap(to_percent)
                st.dataframe(df_monthly_display, use_container_width=True)
            else:
                st.warning("⚠️ 未能找到月度收益表格。请检查Excel中表头是否标准。")

        st.markdown("""
    <hr style="border: none; border-top: 1px solid #ccc;">
    <div style="font-size: 12px; color: #888;">
        <strong>风险提示及免责声明：</strong><br>
        本材料所涵括的信息仅供投资者及其委托代销机构与特定对象的交流研讨，不得用于未经允许的其他任何用途。本材料中所含来源于公开资料的信息，
        本公司对这些信息的准确性和完整性不做任何保证，也不保证所包含的信息及相关建议不会发生任何变更，本公司已力求材料内容的客观、公正，但文中的观点、结论及相关建议仅供参考，不代表任何确定性的判断。
        本材料中所含来源于本公司的任何信息，包括过往业绩、产品分析及预测、产品收益预测和相关建议等，均不代表任何定性判断，不代表产品未来运作的实际效果或可能获得的实际收益，其投资回报可能因市场环境等因素的变化而改变。
        本报告及其内容均为保密信息，未经事先书面同意，本报告不可被复制或分发，本报告的内容亦不可向任何第三者披露。一旦阅读本报告，每一潜在阅读者应被视为已同意此项条款。
        除本页条款外，本材料其他内容和任何表述均属不具有法律约束力的用语，不具有任何法律约束力，不构成法律协议的一部分，不应被视为构成向任何人士发出的要约或要约邀请，也不构成任何承诺。
        本材料所含信息仅供参考，具体以相关法律文件为准。
    </div>
""", unsafe_allow_html=True)

    else:
        st.warning(f"未找到产品文件：{file_name}。请确保Excel文件放置在项目目录中。")
else:
    st.info("该分类下暂无可视化产品，敬请期待更多更新。")
