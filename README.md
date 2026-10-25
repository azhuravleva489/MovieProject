# Movie Dashboard

## Table of Content
[Problem Statement](#problem-statement)

[Data Source](#data-source)

[Tools](#tools)

[Dashboard](#dashboard)

[M Code](#m-code)

[Results and Findings](#results-and-findings)

### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors.
We have a dataset containing movie budgets, box office performance, actors, directors, and genres.

Your task is to build an Excel dashboard that provides insights into this dataset.
The dashboard should helps identify:
- The best-performing actors
- The top movies based on box office metrics
- Director performance
- Genre trends
- Seasonal patterns in movie performance
- Any additional insights that can guide future production decisions

The final dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.

### Data Source
Movie Data : The primary dataset used for this analysis is the "Movie Data.xlsx" file, containing detailed information about each movie's performance (box office and budget), actors, directors and genres. You can download the original datasource here: [Movie Dataset Excel file](https://github.com/user-attachments/files/33923158/Movies_Data.xlsx)

### Tools
- Power Query - I used Power Query for Data Cleaning
- Excel - I used Excel for Data Analysis
- Pivot Tables - for Creating the dashboard and Visualizations

### Data Cleaning and Preparation
- Data loading and inspection.
- Handling errors, missing values.
- Data cleaning and formatting. The excel file after the data cleaning & preparation process can be downloaded here - [Movies_Data_Dashboard.xlsx](https://github.com/user-attachments/files/33923251/Movies_Data_Dashboard.xlsx)

### Dashboard
<img width="1444" height="1106" alt="image" src="https://github.com/user-attachments/assets/c16bf0b8-d0e6-4c53-921e-c684e8409013" />

### M code
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\Redmi\Desktop\this\hw_query\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres" = Table.ExpandTableColumn(#"Merged Queries", "Genres", {"Genre"}, {"Genre"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded Genres",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns",{{"Genre", "Genre_First"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Merged Queries1" = Table.NestedJoin(#"Removed Columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres1" = Table.ExpandTableColumn(#"Merged Queries1", "Genres", {"Genre"}, {"Genre"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Expanded Genres1",{{"Genre", "Genre_Second"}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries2" = Table.NestedJoin(#"Reordered Columns1", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Expanded Directors" = Table.ExpandTableColumn(#"Merged Queries2", "Directors", {"Director"}, {"Director"}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Expanded Directors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns2",{{"Director", "Director_First"}}),
    #"Merged Queries3" = Table.NestedJoin(#"Renamed Columns2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors" = Table.ExpandTableColumn(#"Merged Queries3", "Actors", {"Actor"}, {"Actor"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Expanded Actors",{{"Actor", "Cast_First"}}),
    #"Reordered Columns3" = Table.ReorderColumns(#"Renamed Columns3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director_First", "Cast_First_ID", "Cast_First", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries4" = Table.NestedJoin(#"Reordered Columns3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors1" = Table.ExpandTableColumn(#"Merged Queries4", "Actors", {"Actor"}, {"Actor"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Expanded Actors1",{{"Actor", "Cast_Second"}}),
    #"Reordered Columns4" = Table.ReorderColumns(#"Renamed Columns4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director_First", "Cast_First_ID", "Cast_First", "Cast_Second_ID", "Cast_Second", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries5" = Table.NestedJoin(#"Reordered Columns4", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors2" = Table.ExpandTableColumn(#"Merged Queries5", "Actors", {"Actor"}, {"Actor"}),
    #"Renamed Columns5" = Table.RenameColumns(#"Expanded Actors2",{{"Actor", "Cast_Third"}}),
    #"Reordered Columns5" = Table.ReorderColumns(#"Renamed Columns5",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director_First", "Cast_First_ID", "Cast_First", "Cast_Second_ID", "Cast_Second", "Cast_Third_ID", "Cast_Third", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries6" = Table.NestedJoin(#"Reordered Columns5", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors3" = Table.ExpandTableColumn(#"Merged Queries6", "Actors", {"Actor"}, {"Actor"}),
    #"Renamed Columns6" = Table.RenameColumns(#"Expanded Actors3",{{"Actor", "Cast_Fourth"}}),
    #"Reordered Columns6" = Table.ReorderColumns(#"Renamed Columns6",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director_First", "Cast_First_ID", "Cast_First", "Cast_Second_ID", "Cast_Second", "Cast_Third_ID", "Cast_Third", "Cast_Fourth_ID", "Cast_Fourth", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries7" = Table.NestedJoin(#"Reordered Columns6", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors4" = Table.ExpandTableColumn(#"Merged Queries7", "Actors", {"Actor"}, {"Actor"}),
    #"Renamed Columns7" = Table.RenameColumns(#"Expanded Actors4",{{"Actor", "Cast_Fifth"}}),
    #"Reordered Columns7" = Table.ReorderColumns(#"Renamed Columns7",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First", "Genre_Second_ID", "Genre_Second", "Director_First_ID", "Director_First", "Cast_First_ID", "Cast_First", "Cast_Second_ID", "Cast_Second", "Cast_Third_ID", "Cast_Third", "Cast_Fourth_ID", "Cast_Fourth", "Cast_Fifth_ID", "Cast_Fifth", "Budget ($)", "Box Office Revenue ($)"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Reordered Columns7",{{"Release Date", type date}, {"Budget ($)", type number}}),
    #"Changed Type with Locale" = Table.TransformColumnTypes(#"Changed Type1", {{"Budget ($)", type number}}, "en-US"),
    #"Changed Type with Locale1" = Table.TransformColumnTypes(#"Changed Type with Locale", {{"Box Office Revenue ($)", type number}}, "en-US"),
    #"Added Custom" = Table.AddColumn(#"Changed Type with Locale1", "ROI", each ([#"Box Office Revenue ($)"] - [#"Budget ($)"]) / [#"Budget ($)"]),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"ROI", Percentage.Type}})
in
    #"Changed Type2"
```

### Results and Findings
1. Genre Performance
Action is the highest-grossing genre, generating $28.06B in total box office revenue.
<img width="1028" height="314" alt="image" src="https://github.com/user-attachments/assets/598f4a53-7368-47f2-bc97-2bf9f48880f6" />

Horror has the highest average ROI, making it a promising option for budget-efficient productions.
<img width="1020" height="306" alt="image" src="https://github.com/user-attachments/assets/f8b8468a-d750-4543-95a1-d6ad455560f2" />


2. Release Timing
Action movies perform particularly well in May, June, and July.
<img width="1022" height="548" alt="image" src="https://github.com/user-attachments/assets/349dfa04-e79d-4051-87c3-188cce619c11" />

For horror movies, June and September show the highest average box office revenue.
<img width="1016" height="542" alt="image" src="https://github.com/user-attachments/assets/19e1cb89-650a-4bad-98cc-a9ffcd259f63" />



3. Directors and Actors
Zack Snyder and Tom Cruise are strong candidates for action productions based on their historical box office performance.
<img width="1014" height="300" alt="image" src="https://github.com/user-attachments/assets/50cf45da-c110-4ed2-a13d-74e086b45f7d" />
<img width="866" height="298" alt="image" src="https://github.com/user-attachments/assets/8ed5b833-61dd-4b11-8fd2-3b2825166a18" />

James Wan and Patrick Wilson are promising candidates for horror productions.
<img width="1012" height="308" alt="image" src="https://github.com/user-attachments/assets/117790cb-0338-4023-a85d-f51f0974c0a9" />
<img width="868" height="304" alt="image" src="https://github.com/user-attachments/assets/74bbbbbd-ce82-4bf8-9f9f-c076217b3094" />

