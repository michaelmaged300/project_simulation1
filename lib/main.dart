import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'package:file_picker/file_picker.dart';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      home: ExcelSimulation(),
    );
  }
}

class ExcelSimulation extends StatefulWidget {
  @override
  _ExcelSimulationState createState() => _ExcelSimulationState();
}

class _ExcelSimulationState extends State<ExcelSimulation> {
  String filePath = "";
  List<List<dynamic>> smallTable = [];
  List<List<dynamic>> largeTable = [];

  void pickFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (result != null) {
      String? path = result.files.first.path;
      if (path != null) {
        setState(() {
          filePath = path;
        });
        loadExcel(true); // Load small table by default
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text("Added Successfully")),
        );
      }
    }
  }

  void loadExcel(bool isSmallTable) async {
    if (filePath.isEmpty) {
      showErrorMessage("No file selected. Please choose a file.");
      return;
    }

    try {
      var file = File(filePath);
      if (!file.existsSync()) {
        showErrorMessage("File not found! Check the file path.");
        return;
      }

      var bytes = file.readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      var sheet = excel['Sheet1'];
      if (sheet == null) {
        showErrorMessage("Sheet1 not found in the file.");
        return;
      }

      if (isSmallTable) {
        smallTable = _extractTable(sheet, startRow: 10, endRow: 15, startCol: 1, endCol: 4);
      } else {
        largeTable = _extractTable(sheet, startRow: 18, endRow: 23, startCol: 1, endCol: sheet.maxCols);
      }

      setState(() {});
    } catch (e) {
      showErrorMessage("An error occurred while loading the file.");
    }
  }

  List<List<dynamic>> _extractTable(Sheet sheet, {required int startRow, required int endRow, required int startCol, required int endCol}) {
    List<List<dynamic>> table = [];
    for (int i = startRow - 1; i < endRow; i++) {
      List<dynamic> row = [];
      for (int j = startCol - 1; j < endCol; j++) {
        var cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: j, rowIndex: i));
        row.add(cell.value ?? '');
      }
      table.add(row);
    }
    return table;
  }

  void showErrorMessage(String message) {
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: Text("Error"),
        content: Text(message),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(context).pop(),
            child: Text("OK"),
          ),
        ],
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(centerTitle: true, title: Text("Excel Simulation")),
      body: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          ElevatedButton(
            onPressed: pickFile,
            child: Text("Pick Excel File"),
          ),
          ElevatedButton(
            onPressed: filePath.isEmpty ? null : () => loadExcel(false),
            child: Text("Simulate (Show Large Table)"),
          ),
          SizedBox(height: 20),
          Expanded(
            child: ListView(
              children: [
                if (smallTable.isNotEmpty) _buildTable(smallTable, "Small Table"),
                if (largeTable.isNotEmpty) _buildTable(largeTable, "Large Table"),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildTable(List<List<dynamic>> table, String title) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Text(title, style: TextStyle(fontSize: 20, fontWeight: FontWeight.bold)),
        SingleChildScrollView(
          scrollDirection: Axis.horizontal,
          child: DataTable(
            headingRowColor: MaterialStateColor.resolveWith((states) => Colors.blue),
            columns: table.first
                .map((e) => DataColumn(
              label: Text(
                e.toString(),
                style: TextStyle(fontWeight: FontWeight.bold, color: Colors.white),
              ),
            ))
                .toList(),
            rows: table.skip(1).toList().asMap().entries.map((entry) {
              int index = entry.key;
              List<dynamic> row = entry.value;

              return DataRow(
                color: MaterialStateColor.resolveWith(
                        (states) => index.isOdd ? Colors.blue[50]! : Colors.white),
                cells: row
                    .map((cell) => DataCell(
                  ConstrainedBox(
                    constraints: BoxConstraints(maxWidth: 150),
                    child: Text(
                      cell.toString(),
                      overflow: TextOverflow.ellipsis,
                    ),
                  ),
                ))
                    .toList(),
              );
            }).toList(),
          ),
        ),
      ],
    );
  }
}
