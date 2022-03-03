import os
import pandas as pd
import sys
import xlsxwriter
import math
from PIL import Image
import black


def img_height(img):
    im = Image.open(f"{img}")
    return im.size[1]


def create_img(grand_slam, player, idx, df, x, y):
    ax = df.plot.scatter(x=x, y=y)
    fig = ax.get_figure()
    img_path = f"{player}_{grand_slam}_{idx}.png"
    fig.savefig(img_path)
    return img_path


def currentFuncName(n: int = 0) -> str:

    """
    prints fn name under which it ran

    for current func name, specify 0 or no argument.
    for name of caller of current func, specify 1.
    for name of caller of caller of current func, specify 2. etc.
    """

    print(f"{sys._getframe(n + 1).f_code.co_name} OK")
    return sys._getframe(n + 1).f_code.co_name


def slam_details(idx, grand_slam, player):

    curr_df = df[df["Winner"] == player][
        df.Tournament.str.contains(grand_slam[0])
    ].iloc[[int(idx)]]

    return {
        "idx": str(idx),
        "df": curr_df,
        "img": create_img(grand_slam, player, idx, df=curr_df, x="Year", y="Runner"),
    }


class Excel:
    def __init__(self, player, file):

        self.player = player
        self.file = file
        self.excel_tabs_config = excel_tabs_config

    def build(self):

        for tab in self.excel_tabs_config[self.player].keys():

            ExcelTab(self.player, tab).build()

        print(f"{self.file} built")


class ExcelTab:
    def __init__(self, player, tab):
        self.player = player
        self.tab = tab
        self.tab_config = excel_tabs_config[player][tab]

    def build(self):

        global col_vs_maxW

        y0 = 5

        col_vs_maxW = {x: 0 for x in range(1, 20)}
        for idx in self.tab_config.keys():

            add_df_graph_pair(
                tab=self.tab,
                df=self.tab_config[idx]["df"],
                x0=1,
                y0=y0,
                img=self.tab_config[idx]["img"],
                idx=idx,
                gap=2,
            )

            img_rows = math.ceil(img_height(self.tab_config[idx]["img"]) / 50)

            y_to_add = max(len(self.tab_config[idx]["df"]), img_rows)

            y0 = y0 + y_to_add + 2

        col_vs_maxW = {x: 0 for x in range(1, 20)}


class ExcelDF:

    global max_width_currently

    def __init__(self, **kwargs):

        for key, value in kwargs.items():
            setattr(self, key, value)

        self.start_x = self.x0
        self.start_y = self.y0 + 1
        self.end_x = self.start_x + len(self.df.columns)

        self.end_y = self.start_y + len(self.df)

        self.df_header = f"Results #{int(self.idx)+1}"

    def insert(self):

        dfStyler = self.df.style.set_properties(
            subset=list(self.df.columns[1:]), **{"text-align": "center"}
        )

        dfStyler = self.df.style.set_properties(
            subset=self.df.columns[0], **{"text-align": "left"}
        )
        dfStyler.set_table_styles(
            [dict(selector="th", props=[("text-align", "center")])]
        )

        dfStyler.to_excel(
            writer, sheet_name=self.tab, startcol=self.x0, startrow=self.y0, index=False
        )

        currentFuncName()

    def add_df_heading(self):

        merge_format = workbook.add_format(
            {
                "bold": 1,
                "border": 1,
                "font_color": "red",
                "align": "center",
                "valign": "vcenter",
                "fg_color": "white",
            }
        )
        writer.sheets[f"{self.tab}"].set_tab_color("#93C78A")

        writer.sheets[f"{self.tab}"].merge_range(
            self.y0 - 1,
            self.x0,
            self.y0 - 1,
            self.end_x - 1,
            self.df_header,
            merge_format,
        )

        currentFuncName()

    def auto_width(self):

        global col_vs_maxW

        for column in self.df.columns:

            column = str(column)

            col_idx = (self.df).columns.get_loc(column)

            tab_x = col_idx + self.x0

            column_width = max(self.df[column].astype(str).map(len).max(), len(column))

            col_vs_maxW[tab_x] = max(column_width, col_vs_maxW[tab_x])

            writer.sheets[f"{self.tab}"].set_column(
                tab_x,
                col_idx + self.y0,
                col_vs_maxW[tab_x] if col_idx >= 1 else col_vs_maxW[tab_x] * 0.8,
            )
        currentFuncName()

    def add_table_border(self):

        border_fmt = workbook.add_format({"bottom": 2, "top": 2, "left": 2, "right": 2})
        writer.sheets[f"{self.tab}"].conditional_format(
            xlsxwriter.utility.xl_range(
                self.y0, self.x0, self.end_y - 1, self.end_x - 1
            ),
            {"type": "no_errors", "format": border_fmt},
        )
        currentFuncName()

    def format_text_col(self):

        left = workbook.add_format({"align": "left", "italic": 1})

        condition = {
            "type": "text",
            "criteria": "not containing",
            "value": "impossible_string",
            "format": left,
        }

        writer.sheets[f"{self.tab}"].conditional_format(
            self.start_y, self.start_x, self.end_y, self.start_x, condition
        )

        currentFuncName()

    def format_header(self):

        header_format_config = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "vcenter",
                "fg_color": "#b3e5fc",
                "border": 1,
                "align": "center",
            }
        )

        for col_idx, value in enumerate(self.df.columns.values):

            if "Average" in str(value):
                value = value.replace(" ", "\n")

            writer.sheets[f"{self.tab}"].write(
                self.y0, col_idx + self.x0, value, header_format_config
            )

        writer.sheets[f"{self.tab}"].set_row(self.y0, 29)

        currentFuncName()

    def pretty_build(self):

        print("DF start")
        self.insert()
        self.add_table_border()
        self.format_header()
        self.auto_width()
        self.format_text_col()
        self.add_df_heading()
        print("DF done")


class ExcelChart:
    def __init__(self, **kwargs):

        for key, value in kwargs.items():
            setattr(self, key, value)

    def build(self):

        print("Image start")
        scale_factor = 0.5
        writer.sheets[f"{self.tab}"].insert_image(
            self.y0,
            self.x0 + len(self.df.columns) + self.gap,
            self.img,
            {"x_scale": scale_factor, "y_scale": scale_factor},
        )

        currentFuncName()


def add_df_graph_pair(tab, df, x0, y0, gap, img, idx):

    ExcelDF(
        tab=tab,
        df=df,
        x0=x0,
        y0=y0,
        img=img,
        idx=idx,
    ).pretty_build()

    ExcelChart(tab=tab, df=df, img=img, x0=x0, y0=y0, gap=gap).build()

    currentFuncName()


df = pd.read_csv(r"static/tennis_grandslam_wins.csv")

grand_slams = set("_".join(title.split(" ")[1:]) for title in df.Tournament.unique())

df["Winner"] = [player.replace(" ", "_") for player in df["Winner"]]

excel_tabs_config = {
    player: {
        grand_slam: {
            str(idx): slam_details(idx, grand_slam, player)
            for idx in range(
                len(
                    df[df["Winner"] == player][
                        df.Tournament.str.contains(grand_slam[0])
                    ]
                )
            )
        }
        for grand_slam in grand_slams
    }
    for player in ["Rafael_Nadal", "Roger_Federer", "Novak_Djokovic"]
}

for player in ["Roger_Federer", "Novak_Djokovic", "Rafael_Nadal"]:
    print("Starting for", player, "..")
    file = f"output/{player}_report_{pd.Timestamp.now().year}.xlsx"

    with pd.ExcelWriter(file) as writer:
        workbook = writer.book
        Excel(player, file).build()
