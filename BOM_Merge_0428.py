import os
import sys

import pandas as pd

from PySide6.QtCore import Qt, QCoreApplication, QItemSelectionModel
from PySide6.QtWidgets import (QApplication, QFileDialog, QHeaderView, QHBoxLayout, 
                               QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, 
                               QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
                               QDialog, QDialogButtonBox, QFormLayout, QLabel, QVBoxLayout, 
                               QLineEdit, QLabel)
from PySide6.QtGui import QColor
from PySide6 import QtGui
import string



QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)





def tree_to_dataframe(tree_widget):
    data = []

    def visit_node(node, parent=None, level=0):
        if parent is not None:
            row = [level, parent]
            row.append(node.text(1))  # Add PREFIX

            # Add text and background color for each column
            for j in range(tree_widget.columnCount()):
                bg_color = node.background(j).color()
                if bg_color.isValid():  # Check if the background color is valid
                    if bg_color == QColor.fromRgbF(0, 0, 0, 1):  # Check if the background color is black
                        row.append({'text': node.text(j), 'bg_color': Qt.white})  # Use white as the background color
                    else:
                        row.append({'text': node.text(j), 'bg_color': bg_color})
                else:
                    row.append({'text': node.text(j), 'bg_color': Qt.white})  # Use white as the default background color
                
            row.pop(4)  # Remove CLASS for ITM
            data.append(row)

        for i in range(node.childCount()):
            visit_node(node.child(i), node.text(0), level + 1)

    for i in range(tree_widget.topLevelItemCount()):
        item = tree_widget.topLevelItem(i)
        visit_node(item)

    columns = ["LVL", "PARENT", "PREFIX", "ITM", "ITM_DESC", "QTY", "UOM", "SRC", "PROC", "THREAD", "APE"]
    df = pd.DataFrame(data, columns=columns)
    return df




class MyTreeWidgetItem(QTreeWidgetItem):
    def __init__(self, parent=None):
        super(MyTreeWidgetItem, self).__init__(parent)

class MyTreeWidget(QTreeWidget):
     
    def __init__(self, parent=None):
        super(MyTreeWidget, self).__init__(parent)

        self.setHeaderLabels(["CLASS", "PREFIX", "ITM_DESC", "QTY", "UOM", "SRC", "PROC", "THREAD", "APE"])
        self.setAlternatingRowColors(True)
        self.header().setSectionResizeMode(QHeaderView.Interactive)
        self.itemClicked.connect(self.item_clicked)

    def do_search(self, term):
        clist = self.findItems(term, Qt.MatchContains | Qt.MatchRecursive)

        sel = self.selectionModel()
        for item in clist:
            index = self.indexFromItem(item)
            sel.select(index, QItemSelectionModel.Select)
        
    def set_item_background_color(self, item, color):
        for i in range(self.columnCount()):
            item.setBackground(i, color)

    def item_clicked(self, item, column):
        self.parent().load_item_properties(item)

    
class DataFrameDialog(QDialog):
    def __init__(self, parent=None):
        super(DataFrameDialog, self).__init__(parent)
        
        self.setWindowTitle('BOM Comparison')
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)  # Add maximize button to the dialog

        self.left_table = QTableWidget()
        self.right_table = QTableWidget()

        layout = QVBoxLayout()

        # Create table title labels
        left_table_title = QLabel("OLD BOM")
        right_table_title = QLabel("NEW BOM")

        # Create a horizontal layout for table titles
        title_layout = QHBoxLayout()
        title_layout.addWidget(left_table_title)
        title_layout.addWidget(right_table_title)

        # Add the title layout to the main layout
        layout.addLayout(title_layout)

        table_layout = QHBoxLayout()
        table_layout.addWidget(self.left_table)
        table_layout.addWidget(self.right_table)
        layout.addLayout(table_layout)
        self.setLayout(layout)

    def show_dataframes(self, old_df, new_df):
        self.show_dataframe(self.left_table, old_df)
        self.show_dataframe(self.right_table, new_df)
        self.compare_and_highlight_differences(old_df, new_df)
        self.sync_table_scrollbars()
        self.show()


    def show_dataframe(self, table, df):
        table.setRowCount(df.shape[0])
        table.setColumnCount(df.shape[1])
        table.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                cell_data = df.iat[i, j]
                if isinstance(cell_data, dict):  # Check if cell_data is a dictionary
                    item = QTableWidgetItem(cell_data['text'])
                    item.setBackground(cell_data['bg_color'])
                else:
                    item = QTableWidgetItem(str(cell_data))                    
                table.setItem(i, j, item)


    def compare_and_highlight_differences(self, old_df, new_df):
        # Create a set of row labels for each dataframe
        old_row_labels = set(old_df.index)
        new_row_labels = set(new_df.index)

        # Find missing rows in old_df
        missing_rows = new_row_labels - old_row_labels

        # Insert missing rows into the old_df table with a blue background
        for row_label in missing_rows:
            new_row_index = new_df.index.get_loc(row_label)
            self.left_table.insertRow(new_row_index)
            for j in range(old_df.shape[1]):
                item = QTableWidgetItem("")
                item.setBackground(Qt.yellow)
                self.left_table.setItem(new_row_index, j, item)
                


    def sync_table_scrollbars(self):
        def sync_scrollbars(value):
            self.left_table.verticalScrollBar().setValue(value)
            self.right_table.verticalScrollBar().setValue(value)

        left_scrollbar = self.left_table.verticalScrollBar()
        right_scrollbar = self.right_table.verticalScrollBar()

        left_scrollbar.valueChanged.connect(sync_scrollbars)
        right_scrollbar.valueChanged.connect(sync_scrollbars)




class AddRowDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
 
        self.setWindowTitle("Add New Row")

        vbox = QVBoxLayout()

        form_layout = QFormLayout()
        self.class_edit = QLineEdit()
        self.prefix_edit = QLineEdit()
        self.itm_desc_edit = QLineEdit()
        self.qty_edit = QLineEdit()
        self.uom_edit = QLineEdit()

        form_layout.addRow("Class:", self.class_edit)
        form_layout.addRow("Prefix:", self.prefix_edit)
        form_layout.addRow("ITM_DESC:", self.itm_desc_edit)
        form_layout.addRow("QTY:", self.qty_edit)
        form_layout.addRow("UOM:", self.uom_edit)

        vbox.addLayout(form_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        vbox.addWidget(button_box)

        self.setLayout(vbox)

    def get_row_data(self):
       
        return [self.class_edit.text(), self.prefix_edit.text(), self.itm_desc_edit.text(),
                self.qty_edit.text(), self.uom_edit.text()]
    

    def do_search(self, term):
        clist = self.findItems(term, Qt.MatchContains | Qt.MatchRecursive)

        sel = self.selectionModel()
        for item in clist:
            index = self.indexFromItem(item)
            sel.select(index, QItemSelectionModel.Select)
            
            
    def set_item_background_color(self, item, color):
        for i in range(self.columnCount()):
            item.setBackground(i, color)

    def item_clicked(self, item, column):
        self.parent().load_item_properties(item)
 


class myWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.resize(500, 600)  # 위젯 사이즈

        # List for Pandas Dataframes
        self.df_list = []
        self.root = []

        open_btn = QPushButton('엑셀 파일 열기', self)
        tree_btn = QPushButton('Tree구조 보기', self)
        self.line_edit = QLineEdit()
        self.search_button = QPushButton("Search")
        transform_btn = QPushButton('Transform to Table', self)
        
        tree_add_btn = QPushButton('Add Row', self)
        tree_del_btn = QPushButton('Delete Row(s)', self)
        tree_move_up_btn = QPushButton('Move Up', self)
        tree_move_down_btn = QPushButton('Move Down', self)

        color_alt_row_btn = QPushButton('설변 품번 체크', self)
        color_row_btn = QPushButton('지울 품번 체크', self)        
        
        
        
        
        # 수평 박스 배치
        hbox = QHBoxLayout()
        hbox.addWidget(open_btn)
        hbox.addWidget(tree_btn)
        hbox.addWidget(self.line_edit)
        hbox.addWidget(self.search_button)
        
        hbox.addWidget(tree_add_btn)
        hbox.addWidget(tree_del_btn)
        hbox.addWidget(tree_move_up_btn)
        hbox.addWidget(tree_move_down_btn)        
        hbox.addWidget(color_alt_row_btn)
        hbox.addWidget(color_row_btn)        
        
        
        hbox.addWidget(transform_btn)


        self.table = QTableWidget(self)
        self.qtree = MyTreeWidget(self)

        TableTreeBox = QHBoxLayout()
        TableTreeBox.addWidget(self.table)
        TableTreeBox.addWidget(self.qtree)
 
        # 수직 박스 배치
        vbox = QVBoxLayout()
        vbox.addLayout(hbox)
        vbox.addLayout(TableTreeBox)

        change_name_label = QLabel("Change Name:")
        self.old_name_edit = QLineEdit()
        self.new_name_edit = QLineEdit()
        self.change_name_button = QPushButton("Change Name")

        right_hbox = QHBoxLayout()
        right_hbox.addWidget(change_name_label)
        right_hbox.addWidget(self.old_name_edit)
        right_hbox.addWidget(self.new_name_edit)
        right_hbox.addWidget(self.change_name_button)

        self.property_labels = ["CLASS", "PREFIX", "ITM_DESC", "QTY", "UOM", "SRC", "PROC", "THREAD", "APE"]
        self.property_lineedits = [QLineEdit() for _ in range(len(self.property_labels))]
        properties_layout = QHBoxLayout()
        for label, lineedit in zip(self.property_labels, self.property_lineedits):
            properties_layout.addWidget(QLabel(label))
            properties_layout.addWidget(lineedit)

        self.save_properties_btn = QPushButton("Save Properties")
        self.save_properties_btn.clicked.connect(self.save_item_properties)
        properties_layout.addWidget(self.save_properties_btn)

        right_hbox.addLayout(properties_layout)
        hbox.addLayout(right_hbox)

        self.setLayout(vbox)

        # 시그널 연결
        open_btn.clicked.connect(self.clickOpenBtn)
        tree_btn.clicked.connect(self.clickTreeBtn)
        self.search_button.clicked.connect(self.on_search_button_clicked)
        self.change_name_button.clicked.connect(self.on_change_name_button_clicked)
        
        tree_add_btn.clicked.connect(self.clickTreeAddBtn)
        tree_del_btn.clicked.connect(self.clickTreeDelBtn)
        tree_move_up_btn.clicked.connect(self.clickTreeMoveUpBtn)
        tree_move_down_btn.clicked.connect(self.clickTreeMoveDownBtn)

        color_alt_row_btn.clicked.connect(self.clickColorAltRowBtn)
        color_row_btn.clicked.connect(self.clickColorRowBtn)
        transform_btn.clicked.connect(self.on_transform_button_clicked)




    def clickTreeAddBtn(self):
        selected_item = self.qtree.currentItem()

        if not selected_item:
            return

        dialog = AddRowDialog(self)
        result = dialog.exec()

        if result == QDialog.Accepted:
            new_item_data = dialog.get_row_data()
            new_item = QTreeWidgetItem(new_item_data)

            # Set the background color of the new item to yellow
            yellow_background = QColor(255, 255, 0)
            for i in range(new_item.columnCount()):
                new_item.setBackground(i, yellow_background)

            selected_item.parent().insertChild(selected_item.parent().indexOfChild(selected_item) + 1, new_item)



        new_item_data = ['', '', '', '', '']
        new_item = QTreeWidgetItem(new_item_data)
        # selected_item.addChild(new_item) 

    def clickTreeDelBtn(self):
        selected_items = self.qtree.selectedItems()
        for item in selected_items:
            item.parent().removeChild(item)

    def clickTreeMoveUpBtn(self):
        selected_item = self.qtree.currentItem()
        if not selected_item or not selected_item.parent():
            return

        parent = selected_item.parent()
        index = parent.indexOfChild(selected_item)
        if index == 0:
            return

        parent.removeChild(selected_item)
        parent.insertChild(index - 1, selected_item)
        self.qtree.setCurrentItem(selected_item)

    def clickTreeMoveDownBtn(self):
        selected_item = self.qtree.currentItem()
        if not selected_item or not selected_item.parent():
            return

        parent = selected_item.parent()
        index = parent.indexOfChild(selected_item)
        if index == parent.childCount() - 1:
            return

        parent.removeChild(selected_item)
        parent.insertChild(index + 1, selected_item)
        self.qtree.setCurrentItem(selected_item)

    def clickColorRowBtn(self):
        selected_items = self.qtree.selectedItems()
        gray_background = QColor(192, 192, 192)

        for item in selected_items:
            for i in range(item.columnCount()):
                item.setBackground(i, gray_background)


    def clickColorAltRowBtn(self):
        selected_items = self.qtree.selectedItems()
        sky_blue_background = QColor(135, 206, 235)

        for item in selected_items:
            for i in range(item.columnCount()):
                item.setBackground(i, sky_blue_background)
                
     

    def on_transform_button_clicked(self):
        df = tree_to_dataframe(self.qtree)
        self.df_list.append(df)
        old_df = self.df_list[2]
        new_df = self.df_list[3]
        # self.show_dataframe_on_table(self.df_list[-1])
        # self.show_dataframe_in_popup(self.df_list[-1])
        self.show_dataframes_in_popup(old_df, new_df)

    def show_dataframes_in_popup(self, old_df, new_df):
        dialog = DataFrameDialog(self)
        dialog.show_dataframes(old_df, new_df)
        dialog.exec()


    def load_item_properties(self, item):
        for i in range(len(self.property_labels)):
            self.property_lineedits[i].setText(item.text(i))
        self.current_item = item

    def save_item_properties(self):
        for i in range(len(self.property_labels)):
            self.current_item.setText(i, self.property_lineedits[i].text())



    def increment_last_alpha(self, s):
        last_alpha = s[-1]
        if last_alpha in string.ascii_uppercase:
            index = string.ascii_uppercase.index(last_alpha)
            if index + 1 < len(string.ascii_uppercase):
                s = s[:-1] + string.ascii_uppercase[index + 1]
        return s


    def find_nodes_by_name(self, parent_item, name):
        result = []

        if parent_item is None:
            parent_item = self.qtree.invisibleRootItem()

        for i in range(parent_item.childCount()):
            child_item = parent_item.child(i)
            if child_item.text(0) == name:
                result.append(child_item)
            result.extend(self.find_nodes_by_name(child_item, name))

        return result
        

    def change_node_name(self, old_name, new_name):
        # Find all items by old_name
        items = self.find_nodes_by_name(None, old_name)

        for item in items:
            # Update the item's name
            item.setText(0, new_name)

            # Change the background color to red
            self.qtree.set_item_background_color(item, QtGui.QColor("red"))

            # Change parent node names and background color
            self.change_node_and_ancestors(item.parent())

        
    def change_node_and_ancestors(self, item):
        if item is None:
            return

        # Change the node name
        old_name = item.text(0)
        new_name = self.increment_last_alpha(old_name)

        # Update the item's name
        item.setText(0, new_name)

        # Change the background color to red
        self.qtree.set_item_background_color(item, QtGui.QColor("red"))

        # Move to the next parent
        self.change_node_and_ancestors(item.parent())    

    def update_tree(self):
        self.qtree.clear()
        self.clickTreeBtn()

    def update_table(self):
        self.table.clear()
        self.initTableWidget(2)   
    
    def on_change_name_button_clicked(self):
        old_name = self.old_name_edit.text()
        new_name = self.new_name_edit.text()
        self.change_node_name(old_name, new_name)


    
    def clickOpenBtn(self):
        file_path, ext = QFileDialog.getOpenFileName(self, '파일 열기', os.getcwd(), 'excel file (*.xls *.xlsx)')
        if file_path:
            self.df_list = self.loadData(file_path)
            self.initTableWidget(id=2)
            
            
    def on_search_button_clicked(self):
        text = self.line_edit.text()
        self.qtree.selectionModel().clear()
        self.qtree.do_search(text)
        
    def on_change_name_button_clicked(self):
        old_name = self.old_name_edit.text()
        new_name = self.new_name_edit.text()
        self.change_node_name(old_name, new_name)        

    def clickTreeBtn(self):
        TreeWidget = self.qtree
        
        TreeWidget.setColumnCount(5)
        TreeWidget.setHeaderLabels(["CLASS", "PREFIX", "ITM_DESC", "QTY", "UOM", "SRC", "PROC", "THREAD", "APE"] )
        TreeWidget.setAlternatingRowColors(True)
        TreeWidget.header().setSectionResizeMode(QHeaderView.Interactive)
        
        # Use the new create_tree function instead of anytree
        self.create_tree(self.df_list[2], TreeWidget)

        TreeWidget.expandAll()



    def create_tree(self, df, tree_widget):
        parent_dict = {}

        for index, row in df.iterrows():
            itm = row['ITM']
            parent_itm = row['PARENT']

            itm_data = [itm, row['PREFIX'], row['ITM_DESC'], str(row['QTY']),
                        row['UOM'], row['SRC'], row['PROC'], row['THREAD'], row['APE']]

            if parent_itm not in parent_dict:
                parent_dict[parent_itm] = QTreeWidgetItem(tree_widget, [parent_itm])

            itm_item = QTreeWidgetItem(parent_dict[parent_itm], itm_data)
            parent_dict[itm] = itm_item  
    
        
    def loadData(self, file_name):
        df_list = []        
        with pd.ExcelFile(file_name) as wb:            
            for i, sn in enumerate(wb.sheet_names):              
                try:
                    df = pd.read_excel(wb, sheet_name=sn)
                    if sn == "ECO_BOM": #ECO_BOM시트만 OLD BOM정리
                        df= df.iloc[3:,:11]
                        df.replace('\n','', regex=True, inplace=True)
                        df= df.rename(columns = df.iloc[0])
                        df= df.drop(df.index[0])
                        df.dropna(axis=0, how='all', inplace=True)
                        df= df
                        df= df.reset_index(drop=True)
                        
                except Exception as e:
                    print('File read error:', e)
                else:
                    df = df.fillna(0)
                    df.name = sn
                    df_list.append(df)
        return df_list
    
    
    def initTableWidget(self, id):
        # 테이블 위젯 값 쓰기
        self.table.clear()
        # select dataframe
        df = self.df_list[id]; 

        # table write        
        col = len(df.keys())
        self.table.setColumnCount(col)
        self.table.setHorizontalHeaderLabels(df.keys())
 
        row = len(df.index)
        self.table.setRowCount(row)
        self.writeTableWidget(df, row, col)

              
 
    def writeTableWidget(self, df, row, col):
        
        for r in range(row): 
            for c in range(col):
                item = QTableWidgetItem(str(df.iloc[r][c]))
                self.table.setItem(r, c, item) 
        self.table.resizeColumnsToContents()
 

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    w = myWindow()
    w.show()
    sys.exit(app.exec())   
