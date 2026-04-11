# Copyright (c) 2026 Tran Xuan An
# This file is part of CAD Column Inspector Pro.
# Licensed under the MIT License.

# excel_exporter.py
import pandas as pd


class ExcelExporter:
    @staticmethod
    def export_data(data, file_path):
        """Export data to Excel with multiple sheets"""
        if not data:
            return False, "No data to export!"

        df = pd.DataFrame(data)

        # Group by Type, Size and Layer, add representative 'code' column (take the first code in group)
        grouped = (
            df.groupby(["Type", "Size", "Layer"])
            .agg(
                code=("code", lambda x: x.iloc[0] if len(x) > 0 else ""),
                Count=("Size", "size"),
                IDs=("ID", lambda x: ", ".join(map(str, x))),
                Area_avg=("Area", "mean"),
                Area_min=("Area", "min"),
                Area_max=("Area", "max"),
                Coordinates=(
                    "Coordinate_X",
                    lambda x: ", ".join(
                        [
                            f"({x.iloc[i]},{df.iloc[i]['Coordinate_Y']})"
                            for i in range(len(x))
                        ]
                    ),
                ),
            )
            .reset_index()
        )
        # Move 'code' column to 2nd position (after Type)
        cols = list(grouped.columns)
        if "code" in cols:
            cols.remove("code")
            code_idx = cols.index("Type") + 1
            cols = cols[:code_idx] + ["code"] + cols[code_idx:]
            grouped = grouped[cols]

        # Round area columns
        grouped["Area_avg"] = grouped["Area_avg"].round(2)
        grouped["Area_min"] = grouped["Area_min"].round(2)
        grouped["Area_max"] = grouped["Area_max"].round(2)

        try:
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                # Sheet 1: Grouped by Type, Size and Layer
                grouped.to_excel(writer, sheet_name="Summary_by_Type_Size", index=False)

                # Sheet 2: Detailed data (ensure 'code' column is present and at the first position if exists)
                cols = list(df.columns)
                if "code" in cols:
                    cols = ["code"] + [c for c in cols if c != "code"]
                    df = df[cols]
                df.to_excel(writer, sheet_name="Detailed_Data", index=False)

                # Sheet 3: Statistics
                stats_data = ExcelExporter._create_statistics(df)
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(
                    writer, sheet_name="Statistics", index=False, header=False
                )

            summary = {
                "total_items": len(df),
                "rectangles": len(df[df["Type"].str.contains("Rectangle")]),
                "circles": len(df[df["Type"].str.contains("Circle")]),
                "unique_groups": len(grouped),
            }

            return True, summary

        except Exception as e:
            return False, f"Failed to export: {str(e)}"

    @staticmethod
    def _create_statistics(df):
        """Create statistics data"""
        stats_data = []

        # Overall statistics
        stats_data.append(["=== OVERALL STATISTICS ===", ""])
        stats_data.append(["Total Objects", len(df)])
        stats_data.append(["Rectangles", len(df[df["Type"].str.contains("Rectangle")])])
        stats_data.append(["Circles", len(df[df["Type"].str.contains("Circle")])])
        stats_data.append(["Unique Sizes", df["Size"].nunique()])
        stats_data.append(["", ""])

        # Statistics by Type
        stats_data.append(["=== STATISTICS BY TYPE ===", ""])
        for obj_type in df["Type"].unique():
            type_df = df[df["Type"] == obj_type]
            stats_data.append([f"Type: {obj_type}", ""])
            stats_data.append(["  Count", len(type_df)])
            stats_data.append(["  Area - Min", round(type_df["Area"].min(), 2)])
            stats_data.append(["  Area - Max", round(type_df["Area"].max(), 2)])
            stats_data.append(["  Area - Avg", round(type_df["Area"].mean(), 2)])
            stats_data.append(["", ""])

        # Statistics by Size
        stats_data.append(["=== STATISTICS BY SIZE ===", ""])
        size_stats = (
            df.groupby("Size")
            .agg(
                Count=("Size", "size"),
                Type=("Type", lambda x: ", ".join(x.unique())),
                Area_avg=("Area", "mean"),
            )
            .reset_index()
        )
        size_stats["Area_avg"] = size_stats["Area_avg"].round(2)

        # Convert to list for Excel
        stats_data.append(["Size", "Count", "Types", "Avg Area"])
        for _, row in size_stats.iterrows():
            stats_data.append([row["Size"], row["Count"], row["Type"], row["Area_avg"]])

        return stats_data
