package LCT;
//10062019v2
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Set;
import java.util.logging.Level;

import javax.measure.unit.NonSI;
import javax.measure.unit.SI;
import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.DefaultCellEditor;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JCheckBoxMenuItem;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.JRadioButtonMenuItem;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.KeyStroke;
import javax.swing.SwingConstants;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;
import javax.swing.table.TableModel;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jscience.economics.money.Currency;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.atlantec.objects.ObjectManager;
import org.atlantec.objects.Query;
import org.atlantec.util.CollectionsHelper;
import org.atlantec.util.DateHelper;
import org.atlantec.util.logging.LoggingSystem;
import org.apache.log4j.Logger;
import org.atlantec.binding.POID;
import org.atlantec.binding.erm.ActivityInstance;
import org.atlantec.binding.erm.AnalysisCase;
import org.atlantec.binding.erm.AnalysisScenario;
import org.atlantec.binding.erm.Catalogue;
import org.atlantec.binding.erm.CatalogueItemInstance;
import org.atlantec.binding.erm.CatalogueProperty;
import org.atlantec.binding.erm.CataloguePropertyType;
import org.atlantec.binding.erm.EvaluationResult;
import org.atlantec.binding.erm.KeyValue;
import org.atlantec.binding.erm.MassMeasure;
import org.atlantec.binding.erm.MoneyMeasure;
import org.atlantec.binding.erm.ParameterSet;
import org.atlantec.binding.erm.PowerMeasure;
import org.atlantec.binding.erm.ProcessTemplate;
import org.atlantec.binding.erm.ProductComponent;
import org.atlantec.binding.erm.SpecificConsumptionMeasure;
import org.atlantec.catalogue.CatalogueManager;
import org.atlantec.db.Session;
import org.atlantec.directory.InformationDirectory;
import org.atlantec.jeb.ObjectNotFoundException;


public class Gui16052019 extends JPanel {
	public JCheckBox[] getCheckBox() {
		return checkBox;
	}

	public void setCheckBox(JCheckBox[] checkBox) {
		this.checkBox = checkBox;
	}
	
	public static JTabbedPane getTp() {
			return tp;
	}
	public static void setTp(JTabbedPane tp) {
		Gui16052019.tp = tp;
	}
	public JTabbedPane getTp0() {
		return tp0;
	}
	public void setTp0(JTabbedPane tp0) {
		Gui16052019.tp0 = tp0;
	}
	public static JTabbedPane getTp_plot1() {
		return tp_plot1;
	}
	public static void setTp_plot1(JTabbedPane tp_plot1) {
		Gui16052019.tp_plot1 = tp_plot1;
	}
	public static JTabbedPane getTp_plot2() {
		return tp_plot2;
	}
	public static void setTp_plot2(JTabbedPane tp_plot2) {
		Gui16052019.tp_plot2 = tp_plot2;
	}
	public static JTabbedPane getTp_plot3() {
		return tp_plot3;
	}
	public static void setTp_plot3(JTabbedPane tp_plot3) {
		Gui16052019.tp_plot3 = tp_plot3;
	}
	public static JTabbedPane getTp_plot4() {
		return tp_plot4;
	}
	public static void setTp_plot4(JTabbedPane tp_plot4) {
		Gui16052019.tp_plot4 = tp_plot4;
	}
	public static JTabbedPane getTp_plot5() {
		return tp_plot5;
	}
	public static void setTp_plot5(JTabbedPane tp_plot5) {
		Gui16052019.tp_plot5 = tp_plot5;
	}
	public static JTabbedPane getTp_plot6() {
		return tp_plot6;
	}
	public static void setTp_plot6(JTabbedPane tp_plot6) {
		Gui16052019.tp_plot6 = tp_plot6;
	}
	public static JTabbedPane getTp_plot7() {
		return tp_plot7;
	}
	public static void setTp_plot7(JTabbedPane tp_plot7) {
		Gui16052019.tp_plot7 = tp_plot7;
	}
	public static JFrame getFrame() {
		return Frame;
	}
	public static void setFrame11(JFrame frame) {
		Frame = frame;
	}
	public static JFrame getFrame11() {
		return frame;
	}
	public static void setFrame(JFrame frame) {
		Gui16052019.frame = frame;
	}
	public static JFrame getFrame1() {
		return frame1;
	}
	public static void setFrame1(JFrame frame1) {
		Gui16052019.frame1 = frame1;
	}
	public static JFrame getFrame2() {
		return frame2;
	}
	public static void setFrame2(JFrame frame2) {
		Gui16052019.frame2 = frame2;
	}
	public static JFrame getFrame3() {
		return frame3;
	}
	public static void setFrame3(JFrame frame3) {
		Gui16052019.frame3 = frame3;
	}
	public static JFrame getFrame4() {
		return frame4;
	}
	public static void setFrame4(JFrame frame4) {
		Gui16052019.frame4 = frame4;
	}
	public static JFrame getFrame5() {
		return frame5;
	}
	public static void setFrame5(JFrame frame5) {
		Gui16052019.frame5 = frame5;
	}
	public static JFrame getFrame6() {
		return frame6;
	}
	public static void setFrame6(JFrame frame6) {
		Gui16052019.frame6 = frame6;
	}
	public static JFrame getFrame7() {
		return frame7;
	}
	public static void setFrame7(JFrame frame7) {
		Gui16052019.frame7 = frame7;
	}
	public static JFrame getFrame_r() {
		return frame_r;
	}
	public static void setFrame_r(JFrame frame_r) {
		Gui16052019.frame_r = frame_r;
	}
	public static JFrame getFrame_db() {
		return frame_db;
	}
	public static void setFrame_db(JFrame frame_db) {
		Gui16052019.frame_db = frame_db;
	}
	public static JFrame getFrame_il() {
		return frame_il;
	}
	public static void setFrame_il(JFrame frame_il) {
		Gui16052019.frame_il = frame_il;
	}
	public static JPanel getPanel_m() {
		return panel_m;
	}
	public static void setPanel_m(JPanel panel_m) {
		Gui16052019.panel_m = panel_m;
	}
	public static JPanel getPanel_n() {
		return panel_n;
	}
	public static void setPanel_n(JPanel panel_n) {
		Gui16052019.panel_n = panel_n;
	}
	public static JPanel getPanel_FC() {
		return panel_FC;
	}
	public static void setPanel_FC(JPanel panel_FC) {
		Gui16052019.panel_FC = panel_FC;
	}
	public static JPanel getPanel() {
		return panel;
	}
	public static void setPanel(JPanel panel) {
		Gui16052019.panel = panel;
	}
	public static JPanel getPanel0() {
		return panel0;
	}
	public static void setPanel0(JPanel panel0) {
		Gui16052019.panel0 = panel0;
	}
	public static JPanel getPanel0_0() {
		return panel0_0;
	}
	public static void setPanel0_0(JPanel panel0_0) {
		Gui16052019.panel0_0 = panel0_0;
	}
	public static JPanel getPanel_w() {
		return panel_w;
	}
	public static void setPanel_w(JPanel panel_w) {
		Gui16052019.panel_w = panel_w;
	}
	public static JPanel getPanel1() {
		return panel1;
	}
	public static void setPanel1(JPanel panel1) {
		Gui16052019.panel1 = panel1;
	}
	public static JPanel getPanel1_0() {
		return panel1_0;
	}
	public static void setPanel1_0(JPanel panel1_0) {
		Gui16052019.panel1_0 = panel1_0;
	}
	public static JPanel getPanel1_1() {
		return panel1_1;
	}
	public static void setPanel1_1(JPanel panel1_1) {
		Gui16052019.panel1_1 = panel1_1;
	}
	public static JPanel getPanel1_2() {
		return panel1_2;
	}
	public static void setPanel1_2(JPanel panel1_2) {
		Gui16052019.panel1_2 = panel1_2;
	}
	public static JPanel getPanel2() {
		return panel2;
	}
	public static void setPanel2(JPanel panel2) {
		Gui16052019.panel2 = panel2;
	}
	public static JPanel getPanel2_0() {
		return panel2_0;
	}
	public static void setPanel2_0(JPanel panel2_0) {
		Gui16052019.panel2_0 = panel2_0;
	}
	public static JPanel getPanel3() {
		return panel3;
	}
	public static void setPanel3(JPanel panel3) {
		Gui16052019.panel3 = panel3;
	}
	public static JPanel getPanel3_0() {
		return panel3_0;
	}
	public static void setPanel3_0(JPanel panel3_0) {
		Gui16052019.panel3_0 = panel3_0;
	}
	public static JPanel getPanel4() {
		return panel4;
	}
	public static void setPanel4(JPanel panel4) {
		Gui16052019.panel4 = panel4;
	}
	public static JPanel getPanel4_0() {
		return panel4_0;
	}
	public static void setPanel4_0(JPanel panel4_0) {
		Gui16052019.panel4_0 = panel4_0;
	}
	public static JPanel getPanel5() {
		return panel5;
	}
	public static void setPanel5(JPanel panel5) {
		Gui16052019.panel5 = panel5;
	}
	public static JPanel getPanel6() {
		return panel6;
	}
	public static void setPanel6(JPanel panel6) {
		Gui16052019.panel6 = panel6;
	}
	public static JPanel getTest() {
		return test;
	}
	public static void setTest(JPanel test) {
		Gui16052019.test = test;
	}
	public static JPanel getPanel_chart1() {
		return panel_chart1;
	}
	public static void setPanel_chart1(JPanel panel_chart1) {
		Gui16052019.panel_chart1 = panel_chart1;
	}
	public static JPanel getPanel_chart2() {
		return panel_chart2;
	}
	public static void setPanel_chart2(JPanel panel_chart2) {
		Gui16052019.panel_chart2 = panel_chart2;
	}
	public static JPanel getPanel_chart3() {
		return panel_chart3;
	}
	public static void setPanel_chart3(JPanel panel_chart3) {
		Gui16052019.panel_chart3 = panel_chart3;
	}
	public static JPanel getPanel_chart4() {
		return panel_chart4;
	}
	public static void setPanel_chart4(JPanel panel_chart4) {
		Gui16052019.panel_chart4 = panel_chart4;
	}
	public static JPanel getPanel_chart5() {
		return panel_chart5;
	}
	public static void setPanel_chart5(JPanel panel_chart5) {
		Gui16052019.panel_chart5 = panel_chart5;
	}
	public static JPanel getPanel_chart6() {
		return panel_chart6;
	}
	public static void setPanel_chart6(JPanel panel_chart6) {
		Gui16052019.panel_chart6 = panel_chart6;
	}
	public static JPanel getPanel_chart7() {
		return panel_chart7;
	}
	public static void setPanel_chart7(JPanel panel_chart7) {
		Gui16052019.panel_chart7 = panel_chart7;
	}
	public static JPanel getPanel_chart_r() {
		return panel_chart_r;
	}
	public static void setPanel_chart_r(JPanel panel_chart_r) {
		Gui16052019.panel_chart_r = panel_chart_r;
	}
	public static JPanel getPanel_db() {
		return panel_db;
	}
	public static void setPanel_db(JPanel panel_db) {
		Gui16052019.panel_db = panel_db;
	}
	public static JPanel getPanel_db_1() {
		return panel_db_1; 
	}
	public static void setPanel_db_1(JPanel panel_db_1) {
		Gui16052019.panel_db_1 = panel_db_1;
	}
	public static JPanel getPanel_plot1() {
		return panel_plot1;
	}
	public static void setPanel_plot1(JPanel panel_plot1) {
		Gui16052019.panel_plot1 = panel_plot1;
	}
	public static JPanel getPanel_plot2() {
		return panel_plot2;
	}
	public static void setPanel_plot2(JPanel panel_plot2) {
		Gui16052019.panel_plot2 = panel_plot2;
	}
	public static JPanel getPanel_plot3() {
		return panel_plot3;
	}
	public static void setPanel_plot3(JPanel panel_plot3) {
		Gui16052019.panel_plot3 = panel_plot3;
	}
	public static JPanel getPanel_plot4() {
		return panel_plot4;
	}
	public static void setPanel_plot4(JPanel panel_plot4) {
		Gui16052019.panel_plot4 = panel_plot4;
	}
	public static JPanel getPanel_plot5() {
		return panel_plot5;
	}
	public static void setPanel_plot5(JPanel panel_plot5) {
		Gui16052019.panel_plot5 = panel_plot5;
	}
	public static JPanel getPanel_plot6() {
		return panel_plot6;
	}
	public static void setPanel_plot6(JPanel panel_plot6) {
		Gui16052019.panel_plot6 = panel_plot6;
	}
	public static JPanel getPanel_plot7() {
		return panel_plot7;
	}
	public static void setPanel_plot7(JPanel panel_plot7) {
		Gui16052019.panel_plot7 = panel_plot7;
	}
	public static JPanel getPanel_il() {
		return panel_il;
	}
	public static void setPanel_il(JPanel panel_il) {
		Gui16052019.panel_il = panel_il;
	}
	public static JPanel getPanel_s() {
		return panel_s;
	}
	public static void setPanel_s(JPanel panel_s) {
		Gui16052019.panel_s = panel_s;
	}
	public static JButton getImportData() {
		return importData;
	}
	public static void setImportData(JButton importData) {
		Gui16052019.importData = importData;
	}
	public static JButton getButton_search() {
		return button_search;
	}
	public static void setButton_search(JButton button_search) {
		Gui16052019.button_search = button_search;
	}
	public static JButton getButton_download() {
		return button_download;
	}
	public static void setButton_download(JButton button_download) {
		Gui16052019.button_download = button_download;
	}
	public static Dimension getScreenSize() {
		return screenSize;
	}
	public static void setScreenSize(Dimension screenSize) {
		Gui16052019.screenSize = screenSize;
	}
//	public Dimension getSize() {
//		return size;
//	}
//	public void setSize(Dimension size) {
//		Gui16052019.size = size;
//	}
	public JTextField getField() {
		return field;
	}
	public void setField(JTextField field) {
		this.field = field;
	}
	public JTextField getField_search() {
		return field_search;
	}
	public void setField_search(JTextField field_search) {
		this.field_search = field_search;
	}
	public JTextArea getArea() {
		return area;
	}
	public void setArea(JTextArea area) {
		this.area = area;
	}
	public Font getFont_0() {
		return font_0;
	}
	public void setFont_0(Font font_0) {
		this.font_0 = font_0;
	}
	public Font getFont_1() {
		return font_1;
	}
	public void setFont_1(Font font_1) {
		this.font_1 = font_1;
	}
	public Font getFont_2() {
		return font_2;
	}
	public void setFont_2(Font font_2) {
		this.font_2 = font_2;
	}
	public Date getD() {
		return date;
	}
	public void setD(Date date) {
		Gui16052019.date = date;
	}
	public SimpleDateFormat getDateFormat() {
		return dateFormat;
	}
	public void setDateFormat(SimpleDateFormat dateFormat) {
		Gui16052019.dateFormat = dateFormat;
	}
	public static JMenuBar getMenuBar() {
		return menuBar;
	}
	public static void setMenuBar(JMenuBar menuBar) {
		Gui16052019.menuBar = menuBar;
	}
	public static JMenu getMenu() {
		return menu;
	}
	public static void setMenu(JMenu menu) {
		Gui16052019.menu = menu;
	}
	public static JMenu getSubmenu() {
		return submenu;
	}
	public static void setSubmenu(JMenu submenu) {
		Gui16052019.submenu = submenu;
	}
	public static JMenuItem getMenuItem() {
		return menuItem;
	}
	public static void setMenuItem(JMenuItem menuItem) {
		Gui16052019.menuItem = menuItem;
	}
	public static JRadioButtonMenuItem getRbMenuItem() {
		return rbMenuItem;
	}
	public static void setRbMenuItem(JRadioButtonMenuItem rbMenuItem) {
		Gui16052019.rbMenuItem = rbMenuItem;
	}
	public static JCheckBoxMenuItem getCbMenuItem() {
		return cbMenuItem;
	}
	public static void setCbMenuItem(JCheckBoxMenuItem cbMenuItem) {
		Gui16052019.cbMenuItem = cbMenuItem;
	}
	public static ActionListener getAL_1() {
		return AL_1;
	}
	public static void setAL_1(ActionListener aL_1) {
		AL_1 = aL_1;
	}
	public static ActionListener getAL_2() {
		return AL_2;
	}
	public static void setAL_2(ActionListener aL_2) {
		AL_2 = aL_2;
	}
	public static ActionListener getAL_3() {
		return AL_3;
	}
	public static void setAL_3(ActionListener aL_3) {
		AL_3 = aL_3;
	}
	public String getSelection_Number() {
		return selection_Number;
	}
	public void setSelection_Number(String selection_Number) {
		this.selection_Number = selection_Number;
	}
	public String getSelection_Name() {
		return selection_Name;
	}
	public void setSelection_Name(String selection_Name) {
		this.selection_Name = selection_Name;
	}
	public static String getCwd() {
		return cwd;
	}
	public static void setCwd(String cwd) {
		Gui16052019.cwd = cwd;
	}
	public String[] getPhase() {
		return Phase;
	}
	public void setPhase(String[] phase) {
		Phase = phase;
	}
	public String[] getWelcome() {
		return Welcome;
	}
	public void setWelcome(String[] welcome) {
		Welcome = welcome;
	}
	public String[] getDescription() {
		return Description;
	}
	public void setDescription(String[] description) {
		Description = description;
	}
	public String[] getDesign() {
		return Design;
	}
	public void setDesign(String[] design) {
		Design = design;
	}
	public String[] getC_H() {
		return C_H;
	}
	public void setC_H(String[] c_H) {
		C_H = c_H;
	}
	public String[] getC_M() {
		return C_M;
	}
	public void setC_M(String[] c_M) {
		C_M = c_M;
	}
	public String[] getC_A() {
		return C_A;
	}
	public void setC_A(String[] c_A) {
		C_A = c_A;
	}
	public String[] getOperation() {
		return Operation;
	}
	public void setOperation(String[] operation) {
		Operation = operation;
	}
	public String[] getMaintenance() {
		return Maintenance;
	}
	public void setMaintenance(String[] maintenance) {
		Maintenance = maintenance;
	}
	public String[] getScrapping() {
		return Scrapping;
	}
	public void setScrapping(String[] scrapping) {
		Scrapping = scrapping;
	}
	public ArrayList<String> getQ() {
		return q;
	}
	public void setQ(ArrayList<String> q) {
		this.q = q;
	}
	public int getActivity_length() {
		return activity_length;
	}
	public void setActivity_length(int activity_length) {
		this.activity_length = activity_length;
	}
	public JComboBox<String> getCb() {
		return cb;
	}
	public void setCb(JComboBox<String> cb) {
		this.cb = cb;
	}
	public JComboBox<String>[] getCb_m() {
		return cb_m;
	}
	public void setCb_m(JComboBox<String>[] cb_m) {
		this.cb_m = cb_m;
	}
	public JComboBox<String> getCbtest() {
		return cbtest;
	}
	public void setCbtest(JComboBox<String> cbtest) {
		this.cbtest = cbtest;
	}
	public JTable getTable_db1() {
		return table_db1;
	}
	public void setTable_db1(JTable table_db1) {
		this.table_db1 = table_db1;
	}
	public JTable getTable_db2() {
		return table_db2;
	}
	public void setTable_db2(JTable table_db2) {
		this.table_db2 = table_db2;
	}
	public Object[][] getTableData() {
		return tableData;
	}
	public void setTableData(Object[][] tableData) {
		this.tableData = tableData;
	}
	public TableModel getDTM() {
		return DTM;
	}
	public void setDTM(TableModel dTM) {
		DTM = dTM;
	}
	public int getI() {
		return i;
	}
	public void setI(int i) {
		this.i = i;
	}
	public File getFile() {
		return file;
	}
	public void setFile(File file) {
		this.file = file;
	}
	public File[] getFile_group() {
		return file_group;
	}
	public void setFile_group(File[] file_group) {
		this.file_group = file_group;
	}
	public Workbook getWb_num() {
		return wb_num;
	}
	public void setWb_num(Workbook wb_num) {
		this.wb_num = wb_num;
	}
	public int getData_length() {
		return data_length;
	}
	public void setData_length(int data_length) {
		this.data_length = data_length;
	}
	public Workbook[] getWb() {
		return wb;
	}
	public void setWb(Workbook[] wb) {
		this.wb = wb;
	}
	public Sheet[] getSheet() {
		return sheet;
	}
	public void setSheet(Sheet[] sheet) {
		this.sheet = sheet;
	}
	public Cell[][] getItem0() {
		return item0;
	}
	public void setItem0(Cell[][] item0) {
		this.item0 = item0;
	}
	public Cell[][] getCell0() {
		return cell0;
	}
	public void setCell0(Cell[][] cell0) {
		this.cell0 = cell0;
	}
	public Cell[][] getCell1() {
		return cell1;
	}
	public void setCell1(Cell[][] cell1) {
		this.cell1 = cell1;
	}
	public Cell[][] getCell2() {
		return cell2;
	}
	public void setCell2(Cell[][] cell2) {
		this.cell2 = cell2;
	}
	public Cell[][] getUnit0() {
		return unit0;
	}
	public void setUnit0(Cell[][] unit0) {
		this.unit0 = unit0;
	}
	public String[][] getContent0() {
		return content0;
	}
	public void setContent0(String[][] content0) {
		this.content0 = content0;
	}
	public String[][] getContent1() {
		return content1;
	}
	public void setContent1(String[][] content1) {
		this.content1 = content1;
	}
	public String[][] getContent2() {
		return content2;
	}
	public void setContent2(String[][] content2) {
		this.content2 = content2;
	}
	public String[][] getContent3() {
		return content3;
	}
	public void setContent3(String[][] content3) {
		this.content3 = content3;
	}
	public String[][] getContent4() {
		return content4;
	}
	public void setContent4(String[][] content4) {
		this.content4 = content4;
	}
	public String[] getChoices() {
		return choices;
	}
	public void setChoices(String[] choices) {
		this.choices = choices;
	}
	public String[] getChoices_NAME() {
		return choices_NAME;
	}
	public void setChoices_NAME(String[] choices_NAME) {
		this.choices_NAME = choices_NAME;
	}
	public double[] getA_result() {
		return a_result;
	}
	public void setA_result(double[] a_result) {
		this.a_result = a_result;
	}
	public double[] getB_result() {
		return b_result;
	}
	public void setB_result(double[] b_result) {
		this.b_result = b_result;
	}
	public double[] getC_result() {
		return c_result;
	}
	public void setC_result(double[] c_result) {
		this.c_result = c_result;
	}
	public String[] getData0() {
		return data0;
	}
	public void setData0(String[] data0) {
		this.data0 = data0;
	}
	public String[] getData1() {
		return data1;
	}
	public void setData1(String[] data1) {
		this.data1 = data1;
	}
	public String[] getData2() {
		return data2;
	}
	public void setData2(String[] data2) {
		this.data2 = data2;
	}
	public String[] getData3() {
		return data3;
	}
	public void setData3(String[] data3) {
		this.data3 = data3;
	}
	public String[] getData4() {
		return data4;
	}
	public void setData4(String[] data4) {
		this.data4 = data4;
	}
	public Object[][] getData() {
		return data;
	}
	public void setData(Object[][] data) {
		this.data = data;
	}
	public String[][] getData_m1() {
		return data_m1;
	}
	public void setData_m1(String[][] data_m1) {
		this.data_m1 = data_m1;
	}
	public String[][] getData_m2() {
		return data_m2;
	}
	public void setData_m2(String[][] data_m2) {
		this.data_m2 = data_m2;
	}
	public String[][] getData_m3() {
		return data_m3;
	}
	public void setData_m3(String[][] data_m3) {
		this.data_m3 = data_m3;
	}
	public String[][] getData_m() {
		return data_m;
	}
	public void setData_m(String[][] data_m) {
		this.data_m = data_m;
	}
	public static long getSerialversionuid() {
		return serialVersionUID;
	}
	public static Logger getLog() {
		return log;
	}

	public static Session getSession() {
		return session;
	}

	public static void setSession(Session session) {
		Gui16052019.session = session;
	}

	public static ObjectManager getObjectManager() {
		return objectManager;
	}

	public static void setObjectManager(ObjectManager objectManager) {
		Gui16052019.objectManager = objectManager;
	}

	public DecimalFormat getFormatter() {
		return formatter;
	}

	public void setFormatter(DecimalFormat formatter) {
		this.formatter = formatter;
	}

	public static JTabbedPane getTp0r() {
		return tp0r;
	}

	public static void setTp0r(JTabbedPane tp0r) {
		Gui16052019.tp0r = tp0r;
	}

	public int getResult_size() {
		return result_size;
	}

	public void setResult_size(int result_size) {
		this.result_size = result_size;
	}

	public static JPanel getPanel1r() {
		return panel1r;
	}

	public static void setPanel1r(JPanel panel1r) {
		Gui16052019.panel1r = panel1r;
	}

	public static JPanel getPanel1r_0() {
		return panel1r_0;
	}

	public static void setPanel1r_0(JPanel panel1r_0) {
		Gui16052019.panel1r_0 = panel1r_0;
	}

	public static JPanel getPanel1r_1() {
		return panel1r_1;
	}

	public static void setPanel1r_1(JPanel panel1r_1) {
		Gui16052019.panel1r_1 = panel1r_1;
	}

	public static JPanel getPanel1r_2() {
		return panel1r_2;
	}

	public static void setPanel1r_2(JPanel panel1r_2) {
		Gui16052019.panel1r_2 = panel1r_2;
	}

	public String[] getR_H() {
		return R_H;
	}

	public void setR_H(String[] r_H) {
		R_H = r_H;
	}

	public String[] getR_M() {
		return R_M;
	}

	public void setR_M(String[] r_M) {
		R_M = r_M;
	}

	public String[] getR_A() {
		return R_A;
	}

	public void setR_A(String[] r_A) {
		R_A = r_A;
	}

	public String[][] getData_m0() {
		return data_m0;
	}

	public void setData_m0(String[][] data_m0) {
		this.data_m0 = data_m0;
	}

	public String[][] getData_m4() {
		return data_m4;
	}

	public void setData_m4(String[][] data_m4) {
		this.data_m4 = data_m4;
	}
	/*
	 create a CUI with frame, panels, labels, drop lists, text fields and buttons 
	 create action to read date and calculate and generate results base on data read
	 */
	private static final long serialVersionUID = 1L;
    static private final Logger log = Logger.getLogger(GUIexample.class);
    private static Session session;
    private static ObjectManager objectManager;
    DecimalFormat formatter = new DecimalFormat("#0.000000");
	
//interface	
	private static JTabbedPane tp	= new JTabbedPane();	//main tab panel
	private static JTabbedPane tp0	= new JTabbedPane();	//sub tab panel
	private static JTabbedPane tp0r	= new JTabbedPane();	//sub tab panel
	private static JTabbedPane      tp_plot1    = new JTabbedPane();
	private static JTabbedPane      tp_plot2    = new JTabbedPane();
	protected int result_size =10;    
    private JCheckBox[] checkBox= new JCheckBox[result_size];
    
		private static JTabbedPane      tp_plot3    = new JTabbedPane();
		private static JTabbedPane      tp_plot4    = new JTabbedPane();
		private static JTabbedPane      tp_plot5    = new JTabbedPane();
		private static JTabbedPane      tp_plot6    = new JTabbedPane(); 
		private static JTabbedPane      tp_plot7    = new JTabbedPane();
		private static JFrame Frame, frame, frame1,frame2,frame3,frame4,frame5,frame6,frame7, frame_r,frame_db, frame_il ;
		private static JPanel 	 panel_m, panel_n,panel_FC,	panel, 				
	panel0, panel0_0, 
	panel_w,
	panel1, panel1_0, 
			panel1_1, 
			panel1_2,
	panel1r, panel1r_0, 
			panel1r_1, 
			panel1r_2,			
	panel2, panel2_0, 
	panel3,	panel3_0, 
	panel4,	panel4_0,
	panel5,
	panel6, test,
	panel_chart1, panel_chart2, panel_chart3,panel_chart4,panel_chart5,panel_chart6,panel_chart7, panel_chart_r,panel_db,panel_db_1,
                panel_plot1,panel_plot2,panel_plot3,panel_plot4,panel_plot5,panel_plot6,panel_plot7,
                panel_il,panel_s;									//all panels
	
		private static JButton importData, button_search, button_download;
		private static Dimension screenSize = //new Dimension (1280,800);//800,700);
			Toolkit.getDefaultToolkit().getScreenSize();//new Dimension (300,200); //
//		private static Dimension size = new Dimension (800,200); //frame size
		static int w_mid = (int) (screenSize.getWidth()/2.5);
		static int h_mid = (int) (screenSize.getHeight()/2.5);
		private JTextField field, field_search;						//textfield
		JTextField[] field0;
		private JTextArea area;
		
		int f0=(int) (6*screenSize.getWidth()/1650);
		int f1=(int) (18*screenSize.getWidth()/1650);
		int f2=(int) (24*screenSize.getWidth()/1650);
		int f3=(int) (36*screenSize.getWidth()/1650);
		
		private Font font_0 = new Font("Arial", Font.PLAIN, f0);
		private Font font_1 = new Font("Arial", Font.BOLD,f1);
		private Font font_2 = new Font("Arial", Font.BOLD,f2);
		private Font font_3 = new Font("Arial", Font.PLAIN, f3);
		
		private static Date date = new Date();
		private static SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss") ;	
		
		static String project_name;
		
	private static JMenuBar menuBar;
	private static JMenu menu, submenu;
	private static JMenuItem menuItem;
	private static JRadioButtonMenuItem rbMenuItem;
	private static JCheckBoxMenuItem cbMenuItem;
	private static ActionListener AL_1, AL_2, AL_3,AL_4;
	private String selection_Number ;
	private String selection_Name;
	private static String frame_name;
    private static String cwd = System.getProperty("user.dir");
	private static String urlString;
	private static String projectString;
		
	
//define all the strings based on phases and activities	

	public static ActionListener getAL_4() 
	{
		return AL_4;
	}
	public static void setAL_4(ActionListener aL_4) 
	{
		AL_4 = aL_4;
	}

	private String [] Phase = 		
		{"Welcome","Design", "C_H","C_M","C_A","R_H", "R_M", "R_A", "Operation","Maintenance","Scrapping", "Report","Comparisons"};
	
	private String [] Welcome = 	
		{"Objective", "Impact", "Approach"};
	private String [] Description = 	
		{		"The direct economic costs of production (CAPEX), operation and maintenance costs (OPEX) and "
										+"\n"+ 	"repair and end-of-life costs can be used in combination with environmental impact and risk assessment"
										+"\n"+ 	"to evaluate and compare different designs, maintenance and replacement strategies for ships based on "
										+"\n"+ 	"a through life perspective by quantifying:"
										+"\n"+ 	"  	- Life Cycle Cost Assessment (LCCA)"
										+"\n"+ 	"  	- Life Cycle Assessment (LCA) "
										+"\n"+ 	"  	- Risk Assessment (RA)", 
			
			"The SHIPLYS Life Cycle Tool (ShipLCA) has been developed to help designers during the early stages of the design process. "
					+"\n"+ "The software can be used by shipyards and design offices for the evaluation of their early designs"
					+"\n"+ "with respect to cost, environmental assessment, risk assessment and end-of-life considerations."
					+"\n"+ "The software can also be used for the evaluation of retrofitting projects. "
										};

	//these names should be consistent with Database folder's names 
	private String [] Design = 		{"Project Name","Life Span", "Financial data","Sensitivity level", "Ship total price","Ship Particulars"};
	private String [] C_H = 		{"Hull Cutting", "Hull Bending", "Hull Welding", "Hull Blasting","Hull Coating", "Transportation"};
	private String [] C_M = 		{"Engine (M1)", "Generator (M2)","Auxiliary (M3)", "Main Engine Transportation","Aux. System Transportation", "Electricity"};
	private String [] C_A = 		{"Outfitting Particular", "Outfitting Cutting", "Outfitting Welding", "Outfitting Coating", "Transportation","Not in use"};
	private String [] R_H = 		{"Hull Cutting", "Hull Bending", "Hull Welding", "Hull Blasting","Hull Coating", "Transportation"};
	private String [] R_M = 		{"Engine (M1)", "Auxiliary (M2)", "Scrubber (M3)","Transportation", "Electricity", "Not in use"};
	private String [] R_A = 		{"Outfitting Particular", "Outfitting Cutting", "Outfitting Welding", "Outfitting Coating", "Transportation","Not in use"};
	private String [] Operation = 	{"Operation Profile (M1)", "Operation Profile (M2)","Operation Profile (M3)","Fuel oil", "Lub oil","Transportation"};
	private String [] Maintenance = {"Machinery", "Hull", "Outfitting", "Docking", "Spare Parts","Not in use"}; 
	private String [] Scrapping = 	{"Scrapping","Transportation", "Electricity", "Hull","Machinery","Outfitting"}; 
	
	private ArrayList<String> q = new ArrayList<String>();
	private int activity_length = Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length+Operation.length+Maintenance.length+Scrapping.length;

//combobox	= droplist
	private JComboBox<String> cb;					 
	private JComboBox<String>[] cb_m;	
	private JComboBox<String> cbtest;
	private JTable table_db1 ;
	private JTable table_db2 ;
	private Object[][] tableData;
	private TableModel DTM;
	private int i  = 0;
	

//set up database in a default folder	
	private File file = new File (cwd+"/db/");
	private File[]file_group = file.listFiles();
	private File file_1 = new File (cwd+"/db/");
	private File[]file_group_1 = file.listFiles();
	
	
	private Workbook wb_num = Workbook.getWorkbook(new File(cwd+"/db/Template/Template.xls"));

//set up data length of database	
	private int data_length = wb_num.getSheet(0).getRows() ;
	private Workbook [] wb = new Workbook[file_group.length];
	private Sheet[] sheet = new Sheet[file_group.length];
	private Cell[][]	item0  = new Cell[file_group.length][data_length];
	private Cell[][]	cell0  = new Cell[file_group.length][data_length];
	private Cell[][]	cell1  = new Cell[file_group.length][data_length];
	private Cell[][]	cell2  = new Cell[file_group.length][data_length];
	private Cell[][]	unit0  = new Cell[file_group.length][data_length];
	private String[][]	content0  = new String[file_group.length][data_length];
	private String[][]	content1  = new String[file_group.length][data_length];
	private String[][]	content2  = new String[file_group.length][data_length];
	private String[][]	content3  = new String[file_group.length][data_length];
	private String[][]	content4  = new String[file_group.length][data_length];

	private String[] 	choices = new String[file_group.length];
	private String[] 	choices_NAME = new String[file_group.length];
	
	private  double [] a_result =new double[5];
	private  double [] b_result =new double[5];
	private  double [] c_result =new double[5];

//build a template column and data comes from excel database
	
	private  String[] data0 = new String[data_length];
	private  String[] data1 = new String[data_length];	
	private  String[] data2 = new String[data_length];
	private  String[] data3 = new String[data_length];
	private  String[] data4 = new String[data_length];

//build a template table include previous column	
	private Object[][] 	data;
	
//build a column include the select column for each activity		
	private String[][] data_m0	= new String[data_length][activity_length];
	private String[][] data_m1	= new String[data_length][activity_length];
	private String[][] data_m2	= new String[data_length][activity_length];
	private String[][] data_m3	= new String[data_length][activity_length];
	private String[][] data_m4	= new String[data_length][activity_length];
	private String[][] data_m	= new String[data_length][activity_length];
	
//uploading parameters(should be global)
    double Life_span;
    public double getGWP() {
		return GWP;
	}

	public void setGWP(double gWP) {
		GWP = gWP;
	}

	public double getAP() {
		return AP;
	}

	public void setAP(double aP) {
		AP = aP;
	}

	public double getEP() {
		return EP;
	}

	public void setEP(double eP) {
		EP = eP;
	}

	public double getPOCP() {
		return POCP;
	}

	public void setPOCP(double pOCP) {
		POCP = pOCP;
	}
	double PV;
    double Interest;
    double SL;
    double CoTL;
    double sum;
    double GWP;
    double AP;
    double EP; 
    double POCP; 
    double P_GWP; 
    double P_AP; 
    double P_EP; 
    double P_POCP; 
    double RA; 
    double Total_RA; 
    double Total_CRA;
      
//construction function
	@SuppressWarnings("unchecked")
	private  Gui16052019() throws BiffException, IOException{
		
            
            for(i=0;i<data_length;i++) {
                for(int j=0;j<activity_length;j++){
                    data_m1[i][j]="0";
                    data_m2[i][j]="0";
                    data_m3[i][j]="0";
                }
            }
            
            data_m1[4][5]="1"; data_m2[4][5]="1"; data_m3[4][5]="1";
            data_m1[2][6]="1"; data_m2[2][6]="1"; data_m3[2][6]="1";
            data_m1[2][7]="1"; data_m2[2][7]="1"; data_m3[2][7]="1";
            data_m1[2][8]="1"; data_m2[2][8]="1"; data_m3[2][8]="1";
            data_m1[2][9]="1"; data_m2[2][9]="1"; data_m3[2][9]="1";
            data_m1[2][10]="1"; data_m2[2][10]="1"; data_m3[2][10]="1";
            
            data_m1[1][17]="1"; data_m2[1][17]="1"; data_m3[1][17]="1";
            
            data_m1[2][19]="1"; data_m2[2][19]="1"; data_m3[2][19]="1";
            data_m1[2][20]="1"; data_m2[2][20]="1"; data_m3[2][20]="1";
            data_m1[2][21]="1"; data_m2[2][21]="1"; data_m3[2][21]="1";
            
            data_m1[2][24]="1"; data_m2[2][24]="1"; data_m3[2][24]="1";
            data_m1[2][25]="1"; data_m2[2][25]="1"; data_m3[2][25]="1";
            data_m1[2][26]="1"; data_m2[2][26]="1"; data_m3[2][26]="1";
            data_m1[2][27]="1"; data_m2[2][27]="1"; data_m3[2][27]="1";
            data_m1[2][28]="1"; data_m2[2][28]="1"; data_m3[2][28]="1";

            data_m1[2][37]="1"; data_m2[2][37]="1"; data_m3[2][37]="1";
            data_m1[2][38]="1"; data_m2[2][38]="1"; data_m3[2][38]="1";
            data_m1[2][39]="1"; data_m2[2][39]="1"; data_m3[2][39]="1";

            
            
//icon		
		ImageIcon ii = createImageIcon(cwd+"/pic/icon.png");	

		
//Welcome page
		panel_w = createPanel("");	//add panel 
	     tp.addTab("Welcome",ii,panel_w,"Welcome");
	     add(tp);
	     tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
	     
	     panel_w.setLayout(new GridBagLayout());
	     GridBagConstraints c_w = new GridBagConstraints();
	     JPanel[] j_w = new JPanel[Welcome.length];
	     JTextArea[] area_w = new JTextArea[Welcome.length];
	     c_w.fill = GridBagConstraints.VERTICAL;

	     for (i=0;i<Welcome.length-1;i++){
	    	 j_w[i] = createWelcomePanel(Welcome[i]);
	    	 j_w[i].setName(Welcome[i]);
		        
			    c_w.weighty = 1;
			    c_w.gridy = i;
			    c_w.gridx = 0;
			    panel_w.add(j_w[i], c_w);
				area_w[i]=area;			    
				area_w[i].setText(Description[i]);
				area_w[i].setFont(font_1);
				area_w[i].setSelectedTextColor(Color.WHITE);
				area_w[i].setSelectionColor(Color.RED);
				area_w[i].setWrapStyleWord(true);
				area_w[i].setBackground(Color.WHITE);
				}
			JLabel TP_lbl = new JLabel(Welcome[2]);	
			TP_lbl.setFont(font_2);
			c_w.weightx = 0;
		    c_w.gridy = 2;
		    c_w.gridx = 0;
			panel_w.add(TP_lbl,c_w);
	     	JTextPane TP = new JTextPane();
	     	JScrollPane jsp = new JScrollPane(TP);
			TP.insertIcon(new ImageIcon ( cwd+"/pic/approach2.png" ));
			TP.setLayout(new BorderLayout());
			
			TP.setEditable(false);
//			Dimension D_approach = new Dimension();
//			D_approach.setSize(screenSize.getWidth()/2,screenSize.getHeight()/2);
			
		 	double w = screenSize.getWidth()*915/1680;
		 	double h = screenSize.getHeight()*420/1050;
		 	Dimension d = new Dimension();
		 	d.setSize(w, h);
		 	jsp.setPreferredSize(d);
			TP.setPreferredSize(d);
		    c_w.weightx = 0;
		    c_w.weighty = 0;
		    c_w.gridy = 3;
		    c_w.gridx = 0;
			panel_w.add(jsp,c_w);

		
//Design phase
		 panel0 = createPanel("");	//add panel 
	     tp.addTab("Design",ii,panel0,"Design");
	     add(tp);
	     tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
          
	     panel0.setLayout(new GridBagLayout());
	     GridBagConstraints c = new GridBagConstraints();
	     JPanel[] j0 = new JPanel[Design.length];
	     field0 = new JTextField[Design.length];
	     JComboBox<String>[]  cb0	=	new JComboBox[Design.length];
	     ActionListener[] db0 = new ActionListener[Design.length];
	     JButton[] IL0 = new JButton[Design.length];
	     ActionListener[] il0 = new ActionListener[Design.length];
	     
	     for (i=0;i<Design.length;i++){
	    	 j0[i] = createsubPanel(String.valueOf(i)+". "+Design[i]);
		        c.fill = GridBagConstraints.HORIZONTAL;
			    c.weightx = 2;
			    c.gridx = i%3;
			    c.gridy = Math.round(i/3);
			    if (j0[i].getName()==""){}
			    else {panel0.add(j0[i], c);}
			    field.setText("0");
			    field.setFont(font_2);
			    field0[i] = field;
			    cb0[i] = cb;
//			    IL0[i] = importData;
			    importData.removeActionListener(createImportListener());;
			    importData.setVisible(false);
				field0[0].setText(frame_name);
				field0[0].setEditable(false);
				}    
	     
			cb0[1].setEditable(false);
			cb0[3].setEditable(false);
			cb0[4].setEditable(false);
			cb0[0].setVisible(false);
			cb0[1].setVisible(false);
			cb0[3].setVisible(false);
			cb0[4].setVisible(false);
			field0[2].setVisible(false);
			field0[5].setVisible(false);
//			IL0[0].setVisible(false);
//			IL0[1].setVisible(false);
//			IL0[3].setVisible(false);
//			IL0[4].setVisible(false);
			
			
//Construction	     
	     //C_H phase	    
	     panel1 = createPanel("");	//add panel 
	     tp.addTab("Construction",ii,panel1,"Construction");
	     panel1.add(tp0);
		 tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
	     
	     panel1_0 = createPanel("");
	     tp0.addTab("Hull",ii,panel1_0,"Hull");
	     panel1_0.setLayout(new GridBagLayout());
	     GridBagConstraints c1 = new GridBagConstraints();
	     JPanel[] j1_0 = new JPanel[C_H.length];
	     JTextField[] field1_0 = new JTextField[C_H.length];
	     JComboBox<String>[]  cb1_0	=	new JComboBox[C_H.length];
	     ActionListener[] db1_0 = new ActionListener[C_H.length];
	     JButton[] IL1_0 = new JButton[C_H.length];
	     ActionListener[] il1_0 = new ActionListener[C_H.length];

	     for (i=0;i<C_H.length;i++){
	    	 j1_0[i] = createsubPanel(String.valueOf(Design.length+i)+". "+C_H[i]);
		        c1.fill = GridBagConstraints.HORIZONTAL;
			    c1.weightx = 2;
			    c1.gridx = i%3;
			    c1.gridy = Math.round(i/3);
			    if ("".equals(j1_0[i].getName())){}
			    else {panel1_0.add(j1_0[i], c1);}
			    field1_0[i] = field;
			    field1_0[i].setVisible(false);
//			    IL1_0[i] = importData;
			    importData.removeActionListener(createImportListener());
			    cb1_0[i]=cb;}
	     

	             	     
	     //C_M phase
	     panel1_1 = createPanel("");
	     tp0.addTab("Machinery",ii,panel1_1,"Machinery");
	     panel1_1.setLayout(new GridBagLayout());
	     GridBagConstraints c2 = new GridBagConstraints();
	     JPanel[] j1_1 = new JPanel[C_M.length];
	     JTextField[] field1_1 = new JTextField[C_M.length];
	     JComboBox<String>[]  cb1_1	=	new JComboBox[C_M.length];
	     ActionListener[] db1_1 = new ActionListener[C_M.length];
	     JButton[] IL1_1 = new JButton[C_M.length];
	     ActionListener[] il1_1 = new ActionListener[C_M.length];

	     for (i=0;i<C_M.length;i++){
	    	 j1_1[i] = createsubPanel(String.valueOf(Design.length+C_H.length+i)+". "+C_M[i]);
		        c2.fill = GridBagConstraints.HORIZONTAL;
			    c2.weightx = 2;
			    c2.gridx = i%3;
			    c2.gridy = Math.round(i/3);
			    if ("".equals(j1_1[i].getName())){}
			    else {panel1_1.add(j1_1[i], c2);}
			    field1_1[i] = field;
			    field1_1[i].setVisible(false);

			    cb1_1[i]=cb;
			    if(cb.getName().contains("M1")||cb.getName().contains("M2")||cb.getName().contains("M3"))
			    	{
			    	importData.addActionListener(createImportListener());
			    	}
			    IL1_1[i] = importData;

			    }

	     //C_A phase
	     panel1_2 = createPanel("");
	     tp0.addTab("Outfitting",ii,panel1_2,"Outfitting");
	     panel1_2.setLayout(new GridBagLayout());
	     GridBagConstraints c3 = new GridBagConstraints();
	     JPanel[] j1_2 = new JPanel[C_A.length];
	     JTextField[] field1_2 = new JTextField[C_A.length];
	     JComboBox<String>[]  cb1_2	=	new JComboBox[C_A.length];
	     ActionListener[] db1_2 = new ActionListener[C_A.length];
	     JButton[] IL1_2 = new JButton[C_A.length];
	     ActionListener[] il1_2 = new ActionListener[C_A.length];

	     for (i=0;i<C_A.length;i++){
	    	 j1_2[i] = createsubPanel(String.valueOf(Design.length+C_H.length+C_M.length+i)+". "+C_A[i]);
	    	 j1_2[i].setName(C_A[i]);
		        c3.fill = GridBagConstraints.HORIZONTAL;
			    c3.weightx = 2;
			    c3.gridx = i%3;
			    c3.gridy = Math.round(i/3);
			    if ("".equals(j1_2[i].getName())){}
			    else {panel1_2.add(j1_2[i], c3);}
			    field1_2[i] = field;
			    field1_2[i].setVisible(false);

//			    IL1_2[i] = importData;
			    importData.removeActionListener(createImportListener());
			    cb1_2[i]=cb;}
	     	
	     	j1_2[5].setVisible(false);
//			IL1_2[5].setVisible(false);
//			field1_2[5].setVisible(false);
//Retrofitting
	     //R_H phase	    
	     panel1r = createPanel("");	//add panel 
	     tp.addTab("Retrofitting",ii,panel1r,"Retrofitting");
	     panel1r.add(tp0r);
		 tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
	     
	     panel1r_0 = createPanel("");
	     tp0r.addTab("Hull-Retrofitting",ii,panel1r_0,"Hull-Retrofitting");
	     panel1r_0.setLayout(new GridBagLayout());
	     GridBagConstraints c1r = new GridBagConstraints();
	     JPanel[] j1r_0 = new JPanel[C_H.length];
	     JTextField[] field1r_0 = new JTextField[C_H.length];
	     JComboBox<String>[]  cb1r_0	=	new JComboBox[C_H.length];
	     ActionListener[] db1r_0 = new ActionListener[C_H.length];
	     JButton[] IL1r_0 = new JButton[C_H.length];
	     ActionListener[] il1r_0 = new ActionListener[C_H.length];

	     for (i=0;i<R_H.length;i++){
	    	 j1r_0[i] = createsubPanel(String.valueOf(Design.length+C_H.length+C_M.length+C_A.length+i)+". "+R_H[i]);
	    	 j1r_0[i].setName(R_H[i]);
		        c1r.fill = GridBagConstraints.HORIZONTAL;
			    c1r.weightx = 2;
			    c1r.gridx = i%3;
			    c1r.gridy = Math.round(i/3);
			    if ("".equals(j1r_0[i].getName())){}
			    else {panel1r_0.add(j1r_0[i], c1r);}
			    field1r_0[i] = field;
			    field1r_0[i].setVisible(false);
//			    IL1r_0[i] = importData;
			    importData.removeActionListener(createImportListener());
			    cb1r_0[i]=cb;}
	             	     
//R_M phase
	     panel1r_1 = createPanel("");
	     tp0r.addTab("Machinery-Retrofitting",ii,panel1r_1,"Machinery-Retrofitting");
	     panel1r_1.setLayout(new GridBagLayout());
	     GridBagConstraints c2r = new GridBagConstraints();
	     JPanel[] j1r_1 = new JPanel[C_M.length];
	     JTextField[] field1r_1 = new JTextField[C_M.length];
	     JComboBox<String>[]  cb1r_1	=	new JComboBox[C_M.length];
	     ActionListener[] db1r_1 = new ActionListener[C_M.length];
	     JButton[] IL1r_1 = new JButton[C_M.length];
	     ActionListener[] il1r_1 = new ActionListener[C_M.length];

	     for (i=0;i<R_M.length;i++){
	    	 j1r_1[i] = createsubPanel(String.valueOf(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+i)+". "+R_M[i]);
	    	 j1r_1[i].setName(R_M[i]);
		        c2r.fill = GridBagConstraints.HORIZONTAL;
			    c2r.weightx = 2;
			    c2r.gridx = i%3;
			    c2r.gridy = Math.round(i/3);
			    if ("".equals(j1r_1[i].getName())){}
			    else {panel1r_1.add(j1r_1[i], c2r);}
			    field1r_1[i] = field;
			    field1r_1[i].setVisible(false);
			    cb1r_1[i]=cb;
//			    IL1r_1[i] = importData;
			    importData.removeActionListener(createImportListener());


			    }

	     j1r_1[5].setVisible(false);
//	     IL1r_1[5].setVisible(false);
//	     field1r_1[5].setVisible(false);
	     
//R_A phase
	     panel1r_2 = createPanel("");
	     tp0r.addTab("Outfitting-Retrofitting",ii,panel1r_2,"Outfitting-Retrofitting");
	     panel1r_2.setLayout(new GridBagLayout());
	     GridBagConstraints c3r = new GridBagConstraints();
	     JPanel[] j1r_2 = new JPanel[C_A.length];
	     JTextField[] field1r_2 = new JTextField[C_A.length];
	     JComboBox<String>[]  cb1r_2	=	new JComboBox[C_A.length];
	     ActionListener[] db1r_2 = new ActionListener[C_A.length];
	     JButton[] IL1r_2 = new JButton[C_A.length];
	     ActionListener[] il1r_2 = new ActionListener[C_A.length];

	     for (i=0;i<R_A.length;i++){
	    	 j1r_2[i] = createsubPanel(String.valueOf(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+i)+". "+R_A[i]);
	    	 j1r_2[i].setName(R_A[i]);
		        c3r.fill = GridBagConstraints.HORIZONTAL;
			    c3r.weightx = 2;
			    c3r.gridx = i%3;
			    c3r.gridy = Math.round(i/3);
			    if ("".equals(j1r_2[i].getName())){}
			    else {panel1r_2.add(j1r_2[i], c3r);}
			    field1r_2[i] = field;
			    field1r_2[i].setVisible(false);
//			    IL1r_2[i] = importData;
			    importData.removeActionListener(createImportListener());
			    cb1r_2[i]=cb;}	     
	     
	     j1r_2[5].setVisible(false);
//	     IL1r_2[5].setVisible(false);
//	     field1r_2[5].setVisible(false);
	     
//Operation phase	     
	     panel2 = createPanel("");	//add panel 
		 tp.addTab("Operation",ii,panel2,"Operation");
	     tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
         panel2.setLayout(new GridBagLayout());
	     GridBagConstraints c4 = new GridBagConstraints();
	     JPanel[] j2 = new JPanel[Operation.length];
	     JTextField[] field2 = new JTextField[Operation.length];
	     JComboBox<String>[]  cb2	=	new JComboBox[Operation.length];
	     ActionListener[] db2 = new ActionListener[Operation.length];
	     JButton[] IL2 = new JButton[Operation.length];
	     ActionListener[] il2 = new ActionListener[Operation.length];

	     for (i=0;i<Operation.length;i++){
	    	 j2[i] = createsubPanel(String.valueOf(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length+i)+". "+Operation[i]);
	    	 j2[i].setName(Operation[i]);
		        c4.fill = GridBagConstraints.HORIZONTAL;
			    c4.weightx = 2;
			    c4.gridx = i%3;
			    c4.gridy = Math.round(i/3);
			    if ("".equals(j2[i].getName())){}
			    else {panel2.add(j2[i], c4);}
			    field2[i] = field;
			    field2[i].setVisible(false);
//			    IL2[i] = importData;
			    importData.removeActionListener(createImportListener());
			    cb2[i]=cb;}
	     
//Maintenance phase	    
	     panel3 = createPanel("");	//add panel 
		 tp.addTab("Maintenance",ii,panel3,"Maintenance");
	     tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
	     panel3.setLayout(new GridBagLayout());
	     GridBagConstraints c5 = new GridBagConstraints();
	     JPanel[] j3 = new JPanel[Maintenance.length];
	     JTextField[] field3 = new JTextField[Maintenance.length];
	     JComboBox<String>[]  cb3	=	new JComboBox[Maintenance.length];
	     ActionListener[] db3 = new ActionListener[Maintenance.length];
	     JButton[] IL3 = new JButton[Maintenance.length];
	     ActionListener[] il3 = new ActionListener[Maintenance.length];

	     for (i=0;i<Maintenance.length;i++){
	    	 j3[i] = createsubPanel(String.valueOf(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length+Operation.length+i)+". "+Maintenance[i]);
	    	 j3[i].setName(Maintenance[i]);
		        c5.fill = GridBagConstraints.HORIZONTAL;
			    c5.weightx = 2;
			    c5.gridx = i%3;
			    c5.gridy = Math.round(i/3);
			    if ("".equals(j3[i].getName())){}
			    else {panel3.add(j3[i], c5);}
			    field3[i] = field;
			    field3[i].setVisible(false);
			    cb3[i]=cb;
//			    IL3[i] = importData;
			    importData.removeActionListener(createImportListener());

	     }
	     
	     j3[5].setVisible(false);
//	     IL3[5].setVisible(false);
//	     field3[5].setVisible(false);
	     
//Scrapping phase	    
	     panel4 = createPanel("");	//add panel 
		 tp.addTab("Scrapping",ii,panel4,"Scrapping");
	     tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
	     panel4.setLayout(new GridBagLayout());
	     GridBagConstraints c6 = new GridBagConstraints();
	     JPanel[] j4 = new JPanel[Scrapping.length];
	     JTextField[] field4 = new JTextField[Scrapping.length];
	     JComboBox<String>[]  cb4	=	new JComboBox[Scrapping.length];
	     ActionListener[] db4 = new ActionListener[Scrapping.length];
	     JButton[] IL4 = new JButton[Scrapping.length];
	     ActionListener[] il4 = new ActionListener[Scrapping.length];

	     for (i=0;i<Scrapping.length;i++){
	    	 if(i>=Scrapping.length){
	    		 break;
	    	 }
	    	 j4[i] = createsubPanel(String.valueOf(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length+Operation.length+Maintenance.length+i)+". "+Scrapping[i]);
		      j4[i].setName(Scrapping[i]);
		        c6.fill = GridBagConstraints.HORIZONTAL;
			    c6.weightx = 2;
			    c6.gridx = i%3;
			    c6.gridy = Math.round(i/3);
			    if ("".equals(j4[i].getName())){}
			    else {panel4.add(j4[i], c6);}
			    field4[i] = field;
			    field4[i].setVisible(false);
			    cb4[i]=cb;
//			    IL4[i] = importData;
			    importData.removeActionListener(createImportListener());
}
	     
//list all combobox in a array and add action for each combobox
	     cb_m = new JComboBox[activity_length];
	     String default_cb = "---=";
	     for(int j=0; j<activity_length;j++)
	     {
	    	 
	    	 if(j<Design.length)
	    	 {
	    		 cb_m[j]=cb0[j];
	    		 db0[j]=createCbActionListener(j);
	    	 	 cb0[j].addActionListener(db0[j]);
	    	 	 }
	    	 
	    	 if(Design.length<=j&&j<Design.length+C_H.length)
	    	 {
	    		 cb_m[j]=cb1_0[j-Design.length];
	    	 	 db1_0[j-Design.length]=createCbActionListener(j);
	    	 	 cb1_0[j-Design.length].addActionListener(db1_0[j-Design.length]);
	    	 	 }
	    	 
	    	 if(Design.length+C_H.length<=j&&j<Design.length+C_H.length+C_M.length)
	    	 {
	    		 cb_m[j]=cb1_1[j-Design.length-C_H.length];
	    		 db1_1[j-Design.length-C_H.length]=createCbActionListener(j);
	    	 	 cb1_1[j-Design.length-C_H.length].addActionListener(db1_1[j-Design.length-C_H.length]);
	    	 	 }
	    	 
	    	 if(Design.length+C_H.length+C_M.length<=j&&j<Design.length+C_H.length+C_M.length+C_A.length)
	    	 {
	    		 cb_m[j]=cb1_2[j-Design.length-C_H.length-C_M.length];
	    		 db1_2[j-Design.length-C_H.length-C_M.length]=createCbActionListener(j);
	    	 	 cb1_2[j-Design.length-C_H.length-C_M.length].addActionListener(db1_2[j-Design.length-C_H.length-C_M.length]);
	    	 	 }
	     
	    	 if(Design.length+C_H.length+C_M.length+C_A.length<=j&&j<Design.length+C_H.length+C_M.length+C_A.length+R_H.length)
	    	 {
	    		 cb_m[j]=cb1r_0[j-Design.length-C_H.length-C_M.length-C_A.length];
	    	 	 db1r_0[j-Design.length-C_H.length-C_M.length-C_A.length]=createCbActionListener(j);
	    	 	 cb1r_0[j-Design.length-C_H.length-C_M.length-C_A.length].addActionListener(db1r_0[j-Design.length-C_H.length-C_M.length-C_A.length]);
	    	 	 }
	    	 
	    	 if(Design.length+C_H.length+C_M.length+C_A.length+R_H.length<=j&&j<Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length)
	    	 {
	    		 cb_m[j]=cb1r_1[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length];
	    		 db1r_1[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length]=createCbActionListener(j);
	    	 	 cb1r_1[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length].addActionListener(db1r_1[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length]);
	    	 	 }
	    	 
	    	 if(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length<=j&&j<Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length)
	    	 {
	    		 cb_m[j]=cb1r_2[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length];
	    		 db1r_2[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length]=createCbActionListener(j);
	    	 	 cb1r_2[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length].addActionListener(db1r_2[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length]);
	    	 	 }
	    	 /////////////////
	    	 
	    	 if(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length<=j&&j<Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length+Operation.length)
	    	 {
	    		 cb_m[j]=cb2[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length];
	    		 db2[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length]=createCbActionListener(j);
	    	 	 cb2[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length].addActionListener(db2[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length]);
	    	 	 }
	    	 
	    	 if(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length+Operation.length<=j&&j<Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length+Operation.length+Maintenance.length)
	    	 {
	    		 cb_m[j]=cb3[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length-Operation.length];
	    		 db3[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length-Operation.length]=createCbActionListener(j);
	    		 cb3[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length-Operation.length].addActionListener(db3[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length-Operation.length]);
	    		 }
	    	 
	    	 else if(Design.length+C_H.length+C_M.length+C_A.length+R_H.length+R_M.length+R_A.length+Operation.length+Maintenance.length<=j&&j<activity_length)
	    	 {
	    		 cb_m[j]=cb4[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length-Operation.length-Maintenance.length];
	    		 db4[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length-Operation.length-Maintenance.length]=createCbActionListener(j);
	    		 cb4[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length-Operation.length-Maintenance.length].addActionListener(db4[j-Design.length-C_H.length-C_M.length-C_A.length-R_H.length-R_M.length-R_A.length-Operation.length-Maintenance.length]);
	    		 }
	    	

	    	 
	     }
 	 	
 	 	

// 
//Report	    
	     panel5 = createPanel("");	//add panel 
		 tp.addTab("Report",ii,panel5,"Report");
	     tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
	     panel5.setLayout(new BoxLayout(panel5, BoxLayout.PAGE_AXIS));
	     setBorder(BorderFactory.createEmptyBorder(0, 10, 10, 10));
     
	     JTextPane textReport = new JTextPane();
	     textReport.setAlignmentX(Component.CENTER_ALIGNMENT);
	     textReport.setSelectionColor(Color.red);	
	     textReport.setSelectedTextColor(Color.white);

	     JScrollPane sp = new JScrollPane(textReport);
	     
		 	double w2 = screenSize.getWidth()*800/1680;
		 	double h2 = screenSize.getHeight()*400/1050;
		 	Dimension d2 = new Dimension();
		 	d2.setSize(w2, h2);
	     
	     sp.setPreferredSize(d2);
 
	     panel5.add(sp);
	     
//add button	
	     
	     JButton calcButton = new JButton("Calculate");
	     calcButton.setName("calcButton");

//add button	     
	     JButton uploadButton = new JButton("Upload");
	     uploadButton.setName("uploadButton");
	     
	     JPanel buttonPane = new JPanel();
	     buttonPane.setLayout(new BoxLayout(buttonPane, BoxLayout.LINE_AXIS));
	     buttonPane.setBorder(BorderFactory.createEmptyBorder(0, 10, 10, 10));
	     buttonPane.add(Box.createHorizontalGlue());
	     buttonPane.add(calcButton);
	     buttonPane.add(Box.createRigidArea(new Dimension(10, 0)));
	     buttonPane.add(uploadButton);
	     panel5.add(buttonPane);

//add action for clicking button
	     //calc
	     
	     calcButton.addActionListener(new ActionListener() {
			
			
			public void actionPerformed(ActionEvent e) {
				
//				add save and calculate only function
			
				JFrame F_db_name = new JFrame("Name result");
				JPanel P_db_name = new JPanel();
				
				F_db_name.setLocation(w_mid, h_mid);
				F_db_name.setSize(new Dimension(350,120));
				JTextField TF_db_name = new JTextField("new result name");
				TF_db_name.setFont(font_1);
				TF_db_name.setPreferredSize(new Dimension(200,50));
				JButton save_name = new JButton("Save");
				JButton cancel_name = new JButton("Calculate only");
				save_name.setFont(font_1);
				cancel_name.setFont(font_1);
				save_name.setPreferredSize(new Dimension(100,20));
				cancel_name.setPreferredSize(new Dimension(200,20));
				
				P_db_name.setLayout(new BorderLayout());
				P_db_name.add(TF_db_name,BorderLayout.PAGE_START);
				P_db_name.add(save_name,BorderLayout.WEST);
				P_db_name.add(cancel_name,BorderLayout.EAST);
				
				F_db_name.add(P_db_name);
				F_db_name.setVisible(true);
				F_db_name.setResizable(false); 
				
				save_name.addActionListener(new ActionListener() {
					
					@Override
					public void actionPerformed(ActionEvent e) {
//						cancel_name.getActionListeners();

						
				
                            try { 
                            	
PV =Double.parseDouble(data_m1[1][2]); //0 means not using present value; 1 means using;
Life_span = Double.parseDouble(field0[1].getText()); //Life span of target (year)
Interest= Double.parseDouble(data_m1[2][2]); //Interest rate (100%)
P_GWP = Double.parseDouble(data_m1[3][2]);
P_AP = Double.parseDouble(data_m1[4][2]);

project_name = field0[0].getText();						//project name
SL= Double.parseDouble(field0[3].getText());			//sensitivity level
CoTL = Double.parseDouble(field0[4].getText());		//ship total price
String PV_state, SL_state = null;
//Statement to identify PV/SL
if(PV==0)
{
	PV_state = "Present value is not applied;";
}
else
{
	PV_state = "Present value is applied;";
}

if(SL==0)
{
	data_m=data_m1;
	SL_state="Average value is applied;";
	System.out.println(SL_state);
}


if(SL==1)
{
	data_m=data_m2;
	SL_state="Minimum value is applied;";
	System.out.println(SL_state);
}


if(SL==2)
{
	data_m=data_m3;
	SL_state="Maximum value is applied;";
	System.out.println(SL_state);
}


System.out.println("test"+data_m[1][12]);
//Construction
//System: ME, MG, Boiler...
Parameter_C_System CS1 = new Parameter_C_System();
CS1.	Engine_type	=	data_m[0][12];					//Engine type
CS1.	Number	=	Double.parseDouble(data_m[1][12])	; //Number of items
CS1.	Weight	=	Double.parseDouble(data_m[2][12])	; //Weight of item (ton)
CS1.	Price	=	Double.parseDouble(data_m[3][12])	; //Item price (Euro/ton)
CS1.	Transportation_type = cb_m[15].getName();							//transportation type
CS1.	Transportation_distance	=	Double.parseDouble(data_m[1][15])	; //Distance of item transportation (km)
CS1.	Transportation_fee	=	Double.parseDouble(data_m[2][15])	; //Transportation price per km (Euro/km)
CS1.	Transportation_SFOC	=	Double.parseDouble(data_m[3][15])	; //Transportation fuel consumption (ton/km)
CS1.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][15])	; //Fuel price of transportation (Euro/ton)
CS1.	Electricity_type = cb_m[17].getName(); //Electricity_type
CS1.	Installation_energy_price	=	Double.parseDouble(data_m[1][17])	; //Energy price (Euro/unit)
CS1. 	Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][15])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
CS1. 	Spec_AP_Trans 	=	Double.parseDouble(data_m[6][15])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
CS1. 	Spec_EP_Trans 	=	Double.parseDouble(data_m[7][15])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
CS1. 	Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][15])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
CS1. 	Spec_GWP_E 	=	Double.parseDouble(data_m[2][17])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
CS1. 	Spec_AP_E 	=	Double.parseDouble( data_m[3][17]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
CS1. 	Spec_EP_E 	=	Double.parseDouble( data_m[4][17])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
CS1. 	Spec_POCP_E 	=	Double.parseDouble(data_m[5][17])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
CS1.H = 7;
CS1.F[0]=Double.parseDouble(data_m1[10][12]);
CS1.F[1]=Double.parseDouble(data_m1[11][12]);
CS1.F[2]=Double.parseDouble(data_m1[12][12]);
CS1.F[3]=Double.parseDouble(data_m1[13][12]);
CS1.F[4]=Double.parseDouble(data_m1[14][12]);
CS1.F[5]=Double.parseDouble(data_m1[15][12]);
CS1.F[6]=Double.parseDouble(data_m1[16][12]);
CS1.C[0]=Double.parseDouble(data_m2[10][12]);
CS1.C[1]=Double.parseDouble(data_m2[11][12]);
CS1.C[2]=Double.parseDouble(data_m2[12][12]);
CS1.C[3]=Double.parseDouble(data_m2[13][12]);
CS1.C[4]=Double.parseDouble(data_m2[14][12]);
CS1.C[5]=Double.parseDouble(data_m2[15][12]);
CS1.C[6]=Double.parseDouble(data_m2[16][12]);
CS1.M[0]=Double.parseDouble(data_m3[10][12]);
CS1.M[1]=Double.parseDouble(data_m3[11][12]);
CS1.M[2]=Double.parseDouble(data_m3[12][12]);
CS1.M[3]=Double.parseDouble(data_m3[13][12]);
CS1.M[4]=Double.parseDouble(data_m3[14][12]);
CS1.M[5]=Double.parseDouble(data_m3[15][12]);
CS1.M[6]=Double.parseDouble(data_m3[16][12]);
CS1.CoTL =CoTL;
CS1.run(); //Run the calculation
Parameter_C_System CS2 = new Parameter_C_System();
CS2.	Engine_type = data_m[0][14];					//battery name
CS2.	Number	=	Double.parseDouble(data_m[1][14])	; //Number of items
CS2.	Weight	=	Double.parseDouble(data_m[2][14])	; //Weight of item (ton)
CS2.	Price	=	Double.parseDouble(data_m[3][14])	; //Item price (Euro/ton)
CS2.	Transportation_type = data_m[0][16];							//transportation type
CS2.	Transportation_distance	=	Double.parseDouble(data_m[1][16])	; //Distance of item transportation (km)
CS2.	Transportation_fee	=	Double.parseDouble(data_m[2][16])	; //Transportation fee per km (Euro/km)
CS2.	Transportation_SFOC	=	Double.parseDouble(data_m[3][16])	; //Transportation fuel consumption (ton/km)
CS2.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][16])	; //Fuel price of transportation (Euro/ton)
CS2.	Electricity_type = data_m[0][17]; //Electricity_type
CS2.	Installation_energy_price	=	Double.parseDouble(data_m[1][17])	; //Energy price (Euro/unit)
CS2. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][16])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
CS2. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][16])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
CS2. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][16])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
CS2. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][16])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
CS2. Spec_GWP_E 	=	Double.parseDouble(data_m[2][17])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
CS2. Spec_AP_E 	=	Double.parseDouble( data_m[3][17]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
CS2. Spec_EP_E 	=	Double.parseDouble( data_m[4][17])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
CS2. Spec_POCP_E 	=	Double.parseDouble(data_m[5][17])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
CS2.H = 7;
CS2.F[0]=Double.parseDouble(data_m1[12][14]);
CS2.F[1]=Double.parseDouble(data_m1[13][14]);
CS2.F[2]=Double.parseDouble(data_m1[14][14]);
CS2.F[3]=Double.parseDouble(data_m1[15][14]);
CS2.F[4]=Double.parseDouble(data_m1[16][14]);
CS2.F[5]=Double.parseDouble(data_m1[17][14]);
CS2.F[6]=Double.parseDouble(data_m1[18][14]);
CS2.C[0]=Double.parseDouble(data_m2[12][14]);
CS2.C[1]=Double.parseDouble(data_m2[13][14]);
CS2.C[2]=Double.parseDouble(data_m2[14][14]);
CS2.C[3]=Double.parseDouble(data_m2[15][14]);
CS2.C[4]=Double.parseDouble(data_m2[16][14]);
CS2.C[5]=Double.parseDouble(data_m2[17][14]);
CS2.C[6]=Double.parseDouble(data_m2[18][14]);
CS2.M[0]=Double.parseDouble(data_m3[12][14]);
CS2.M[1]=Double.parseDouble(data_m3[13][14]);
CS2.M[2]=Double.parseDouble(data_m3[14][14]);
CS2.M[3]=Double.parseDouble(data_m3[15][14]);
CS2.M[4]=Double.parseDouble(data_m3[16][14]);
CS2.M[5]=Double.parseDouble(data_m3[17][14]);
CS2.M[6]=Double.parseDouble(data_m3[18][14]);
CS2.CoTL =CoTL;
CS2.run(); //Run the calculation


Parameter_C_System CS3 = new Parameter_C_System();
CS3.	Engine_type = data_m[0][13];					//Genset name
CS3.	Number	=	Double.parseDouble(data_m[1][13])	; //Number of items
CS3.	Weight	=	Double.parseDouble(data_m[2][13])	; //Weight of item (ton)
CS3.	Price	=	Double.parseDouble(data_m[3][13])	; //Item price (Euro/ton)
CS3.	Transportation_type = data_m[0][15];							//transportation type
CS3.	Transportation_distance	=	Double.parseDouble(data_m[1][15])	; //Distance of item transportation (km)
CS3.	Transportation_fee	=	Double.parseDouble(data_m[2][15])	; //Transportation fee per km (Euro/km)
CS3.	Transportation_SFOC	=	Double.parseDouble(data_m[3][15])	; //Transportation fuel consumption (ton/km)
CS3.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][15])	; //Fuel price of transportation (Euro/ton)
CS3.	Electricity_type = data_m[0][17]; //Electricity_type
CS3.	Installation_energy_price	=	Double.parseDouble(data_m[1][17])	; //Energy price (Euro/unit)
CS3. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][15])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
CS3. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][15])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
CS3. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][15])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
CS3. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][15])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
CS3. Spec_GWP_E 	=	Double.parseDouble(data_m[2][17])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
CS3. Spec_AP_E 	=	Double.parseDouble( data_m[3][17]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
CS3. Spec_EP_E 	=	Double.parseDouble( data_m[4][17])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
CS3. Spec_POCP_E 	=	Double.parseDouble(data_m[5][17])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
CS3.H = 7;
CS3.F[0]=Double.parseDouble(data_m1[10][13]);
CS3.F[1]=Double.parseDouble(data_m1[11][13]);
CS3.F[2]=Double.parseDouble(data_m1[12][13]);
CS3.F[3]=Double.parseDouble(data_m1[13][13]);
CS3.F[4]=Double.parseDouble(data_m1[14][13]);
CS3.F[5]=Double.parseDouble(data_m1[15][13]);
CS3.F[6]=Double.parseDouble(data_m1[16][13]);
CS3.C[0]=Double.parseDouble(data_m2[10][13]);
CS3.C[1]=Double.parseDouble(data_m2[11][13]);
CS3.C[2]=Double.parseDouble(data_m2[12][13]);
CS3.C[3]=Double.parseDouble(data_m2[13][13]);
CS3.C[4]=Double.parseDouble(data_m2[14][13]);
CS3.C[5]=Double.parseDouble(data_m2[15][13]);
CS3.C[6]=Double.parseDouble(data_m2[16][13]);
CS3.M[0]=Double.parseDouble(data_m3[10][13]);
CS3.M[1]=Double.parseDouble(data_m3[11][13]);
CS3.M[2]=Double.parseDouble(data_m3[12][13]);
CS3.M[3]=Double.parseDouble(data_m3[13][13]);
CS3.M[4]=Double.parseDouble(data_m3[14][13]);
CS3.M[5]=Double.parseDouble(data_m3[15][13]);
CS3.M[6]=Double.parseDouble(data_m3[16][13]);
CS3.CoTL =CoTL;
CS3.run(); //Run the calculation

//Material: steel, aluminium, composite material...
Parameter_C_Material CM1 = new Parameter_C_Material();
//

CM1.	LOA	=Double.parseDouble(data_m[	1	][5]);
CM1.	LBP	=Double.parseDouble(data_m[	2	][5]);
CM1.	B	=Double.parseDouble(data_m[	3	][5]);	
CM1.	D	=Double.parseDouble(data_m[	4	][5]);
CM1.	T	=Double.parseDouble(data_m[	5	][5]);
CM1.	Cb	=Double.parseDouble(data_m[	6	][5]);
CM1.	Vs	=Double.parseDouble(data_m[	7	][5]);
CM1.	NJ	=Double.parseDouble(data_m[	8	][5]);
CM1.	NE	=Double.parseDouble(data_m[	9	][5]);
CM1.	kA	=Double.parseDouble(data_m[	10	][5]);
CM1.	Pw	=Double.parseDouble(data_m[	11	][5]);

if(Double.parseDouble(data_m[	12	][5])!=0) 
{
	CM1. 	W11 = Double.parseDouble(data_m[	12	][5]);
}

CM1.	E_price		= Double.parseDouble(data_m[	1	][17]);
CM1.	Spec_GWP_E 	= Double.parseDouble(data_m[	2	][17]);
CM1.	Spec_AP_E 	= Double.parseDouble(data_m[	3	][17]);
CM1.	Spec_EP_E 	= Double.parseDouble(data_m[	4	][17]);
CM1.	Spec_POCP_E = Double.parseDouble(data_m[	5	][17]);

CM1.	P_Cutting	=Double.parseDouble(data_m[	5	][6]);
CM1.	V_Cutting	=Double.parseDouble(data_m[	2	][6]);
CM1.	L_Cutting	=Double.parseDouble(data_m[	1	][6]);
CM1.	SMC_Cutting	=Double.parseDouble(data_m[	3	][6]);
CM1.	MP_Cutting 	= Double.parseDouble(data_m[	4	][6]);
//CM1.	H_Cutting_Hull	= Double.parseDouble(data_m[	7	][6]);
if(Double.parseDouble(data_m[	7	][6])!=0) 
{
	CM1.	H_Cutting_Hull	= Double.parseDouble(data_m[	7	][6]);
}
else 
{
	CM1.	r_cutting = Double.parseDouble(data_m[	8	][6]);
}

CM1.P_Bending=Double.parseDouble(data_m[	2	][7]);
CM1.V_Bending=Double.parseDouble(data_m[	1	][7]);

if(Double.parseDouble(data_m[	4	][7])!=0) 
{
	CM1.H_Bending_Hull= Double.parseDouble(data_m[	4	][7]);
}
else 
{
	CM1.	r_bending = Double.parseDouble(data_m[	5	][7]);
}

CM1.P_Welding=Double.parseDouble(data_m[	5	][8]);
CM1.V_Welding=Double.parseDouble(data_m[	2	][8]);
CM1.L_Welding=Double.parseDouble(data_m[	1	][8]);
CM1.SMC_Welding=Double.parseDouble(data_m[	3	][8]);
CM1.MP_Welding = Double.parseDouble(data_m[	4	][8]);

if(Double.parseDouble(data_m[	7	][8])!=0) 
{
	CM1.H_Welding_Hull= Double.parseDouble(data_m[	7	][8]);
}
else 
{
	CM1.	r_welding = Double.parseDouble(data_m[	8	][8]);
}

CM1.P_Blasting=Double.parseDouble(data_m[	2	][9]);
CM1.V_Blasting=Double.parseDouble(data_m[	1	][9]);
if(Double.parseDouble(data_m[	4	][9])!=0) 
{
	CM1.H_Blasting_Hull= Double.parseDouble(data_m[	4	][9]);
}
else 
{
	CM1.	r_blasting = Double.parseDouble(data_m[	5	][9]);
}
CM1.P_Painting=Double.parseDouble(data_m[	5	][10]);
CM1.V_Painting=Double.parseDouble(data_m[	2	][10]);
CM1.SMC_Painting=Double.parseDouble(data_m[	3	][10]);
CM1.MP_Painting = Double.parseDouble(data_m[	4	][10]);
if(Double.parseDouble(data_m[	7	][10])!=0) 
{
	CM1.H_Painting_Hull= Double.parseDouble(data_m[	7	][10]);
}
else 
{
	CM1.	r_painting = Double.parseDouble(data_m[	8	][10]);
}

CM1.Transportation_distance=Double.parseDouble(data_m[	1	][11]);
CM1.Transportation_price = Double.parseDouble(data_m[	2	][11]);
CM1.Transportation_SFOC = Double.parseDouble(data_m[	3	][11]);
CM1.Transportation_fuel_price = Double.parseDouble(data_m[	4	][11]);
CM1.Spec_GWP_Trans=Double.parseDouble(data_m[	5	][11]);
CM1.Spec_AP_Trans=Double.parseDouble(data_m[	6	][11]);
CM1.Spec_EP_Trans=Double.parseDouble(data_m[	7	][11]);
CM1.Spec_POCP_Trans=Double.parseDouble(data_m[	8	][11]);

CM1. Labour_fee_cutting =Double.parseDouble(data_m[	6	][6]);
CM1. Labour_fee_bending =Double.parseDouble(data_m[	3	][7]);
CM1. Labour_fee_welding =Double.parseDouble(data_m[	6	][8]);
CM1. Labour_fee_blasting =Double.parseDouble(data_m[	3	][9]);
CM1. Labour_fee_painting =Double.parseDouble(data_m[	6	][10]);

CM1.run(); //Run the calculation


Parameter_C_Material CM2 = new Parameter_C_Material();
//
CM2.	LOA	=Double.parseDouble(data_m[	1	][5]);
CM2.	LBP	=Double.parseDouble(data_m[	2	][5]);
CM2.	B	=Double.parseDouble(data_m[	3	][5]);	
CM2.	D	=Double.parseDouble(data_m[	4	][5]);
CM2.	T	=Double.parseDouble(data_m[	5	][5]);
CM2.	Cb	=Double.parseDouble(data_m[	6	][5]);
CM2.	Vs	=Double.parseDouble(data_m[	7	][5]);
CM2.	NJ	=Double.parseDouble(data_m[	8	][5]);
CM2.	NE	=Double.parseDouble(data_m[	9	][5]);
CM2.	Pw	=Double.parseDouble(data_m[	11	][5]);

CM2.kB=Double.parseDouble(data_m[	1	][18]);
CM2.WB=Double.parseDouble(data_m[	2	][18]);
CM2.E_price=Double.parseDouble(data_m[	1	][17]);
CM2.Spec_GWP_E = Double.parseDouble(data_m[	2	][17]);
CM2.Spec_AP_E = Double.parseDouble(data_m[	3	][17]);
CM2.Spec_EP_E = Double.parseDouble(data_m[	4	][17]);
CM2.Spec_POCP_E = Double.parseDouble(data_m[	5	][17]);

CM2.P_Cutting=Double.parseDouble(data_m[	5	][19]);
CM2.V_Cutting=Double.parseDouble(data_m[	2	][19]);
CM2.L_Cutting=Double.parseDouble(data_m[	1	][19]);
CM2.SMC_Cutting=Double.parseDouble(data_m[	3	][19]);
CM2.MP_Cutting = Double.parseDouble(data_m[	4	][19]);
if(Double.parseDouble(data_m[	7	][19])!=0) 
{
	CM2.H_Cutting_Hull= Double.parseDouble(data_m[	7	][19]);
}
else 
{
	CM2.	r_cutting = Double.parseDouble(data_m[	8	][19]);
}
CM2.P_Welding=Double.parseDouble(data_m[	5	][20]);
CM2.V_Welding=Double.parseDouble(data_m[	2	][20]);
CM2.L_Welding=Double.parseDouble(data_m[	1	][20]);
CM2.SMC_Welding=Double.parseDouble(data_m[	3	][20]);
CM2.MP_Welding = Double.parseDouble(data_m[	4	][20]);
if(Double.parseDouble(data_m[	7	][20])!=0) 
{
	CM2.H_Welding_Hull= Double.parseDouble(data_m[	7	][20]);
}
else {
	CM2.	r_welding = Double.parseDouble(data_m[	8	][20]);
}

CM2.P_Painting=Double.parseDouble(data_m[	5	][21]);
CM2.V_Painting=Double.parseDouble(data_m[	2	][21]);
CM2.L_Painting=Double.parseDouble(data_m[	1	][21]);
CM2.SMC_Painting=Double.parseDouble(data_m[	3	][21]);
CM2.MP_Painting = Double.parseDouble(data_m[	4	][21]);
if(Double.parseDouble(data_m[	7	][21])!=0) 
{
	CM2.H_Painting_Hull= Double.parseDouble(data_m[	7	][21]);
}
else 
{
	CM2.	r_painting = Double.parseDouble(data_m[	8	][21]);
}

CM2. Labour_fee_cutting =Double.parseDouble(data_m[	6	][19]);
CM2. Labour_fee_welding =Double.parseDouble(data_m[	6	][20]);
CM2. Labour_fee_painting =Double.parseDouble(data_m[	6	][21]);

CM2.run(); //Run the calculation

//Retrofitting
//System: ME, Aux, Scrubber...
Parameter_C_System RS1 = new Parameter_C_System();
RS1.	Engine_type	=	data_m[0][30];					//Engine type
RS1.	Number	=	Double.parseDouble(data_m[1][30])	; //Number of items
RS1.	Weight	=	Double.parseDouble(data_m[2][30])	; //Weight of item (ton)
RS1.	Price	=	Double.parseDouble(data_m[3][30])	; //Item price (Euro/ton)
RS1.	Transportation_type = cb_m[32].getName();							//transportation type
RS1.	Transportation_distance	=	Double.parseDouble(data_m[1][32])	; //Distance of item transportation (km)
RS1.	Transportation_fee	=	Double.parseDouble(data_m[2][32])	; //Transportation price per km (Euro/km)
RS1.	Transportation_SFOC	=	Double.parseDouble(data_m[3][32])	; //Transportation fuel consumption (ton/km)
RS1.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][32])	; //Fuel price of transportation (Euro/ton)
RS1.	Electricity_type = cb_m[34].getName(); //Electricity_type
RS1.	Installation_energy_price	=	Double.parseDouble(data_m[1][34])	; //Energy price (Euro/unit)
RS1. 	Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][32])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
RS1. 	Spec_AP_Trans 	=	Double.parseDouble(data_m[6][32])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
RS1. 	Spec_EP_Trans 	=	Double.parseDouble(data_m[7][32])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
RS1. 	Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][32])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
RS1. 	Spec_GWP_E 	=	Double.parseDouble(data_m[2][34])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
RS1. 	Spec_AP_E 	=	Double.parseDouble( data_m[3][34]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
RS1. 	Spec_EP_E 	=	Double.parseDouble( data_m[4][34])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
RS1. 	Spec_POCP_E 	=	Double.parseDouble(data_m[5][34])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
RS1.H = 7;
RS1.F[0]=Double.parseDouble(data_m1[10][30]);
RS1.F[1]=Double.parseDouble(data_m1[11][30]);
RS1.F[2]=Double.parseDouble(data_m1[12][30]);
RS1.F[3]=Double.parseDouble(data_m1[13][30]);
RS1.F[4]=Double.parseDouble(data_m1[14][30]);
RS1.F[5]=Double.parseDouble(data_m1[15][30]);
RS1.F[6]=Double.parseDouble(data_m1[16][30]);
RS1.C[0]=Double.parseDouble(data_m2[10][30]);
RS1.C[1]=Double.parseDouble(data_m2[11][30]);
RS1.C[2]=Double.parseDouble(data_m2[12][30]);
RS1.C[3]=Double.parseDouble(data_m2[13][30]);
RS1.C[4]=Double.parseDouble(data_m2[14][30]);
RS1.C[5]=Double.parseDouble(data_m2[15][30]);
RS1.C[6]=Double.parseDouble(data_m2[16][30]);
RS1.M[0]=Double.parseDouble(data_m3[10][30]);
RS1.M[1]=Double.parseDouble(data_m3[11][30]);
RS1.M[2]=Double.parseDouble(data_m3[12][30]);
RS1.M[3]=Double.parseDouble(data_m3[13][30]);
RS1.M[4]=Double.parseDouble(data_m3[14][30]);
RS1.M[5]=Double.parseDouble(data_m3[15][30]);
RS1.M[6]=Double.parseDouble(data_m3[16][30]);
RS1.CoTL =CoTL;
RS1.run(); //Run the calculation


Parameter_C_System RS2 = new Parameter_C_System();
RS2.	Engine_type = data_m[0][31];					//battery name
RS2.	Number	=	Double.parseDouble(data_m[1][31])	; //Number of items
RS2.	Weight	=	Double.parseDouble(data_m[2][31])	; //Weight of item (ton)
RS2.	Price	=	Double.parseDouble(data_m[3][31])	; //Item price (Euro/ton)
RS2.	Transportation_type = data_m[0][33];							//transportation type
RS2.	Transportation_distance	=	Double.parseDouble(data_m[1][33])	; //Distance of item transportation (km)
RS2.	Transportation_fee	=	Double.parseDouble(data_m[2][33])	; //Transportation fee per km (Euro/km)
RS2.	Transportation_SFOC	=	Double.parseDouble(data_m[3][33])	; //Transportation fuel consumption (ton/km)
RS2.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][33])	; //Fuel price of transportation (Euro/ton)
RS2.	Electricity_type = data_m[0][34]; //Electricity_type
RS2.	Installation_energy_price	=	Double.parseDouble(data_m[1][34])	; //Energy price (Euro/unit)
RS2. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][33])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
RS2. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][33])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
RS2. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][33])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
RS2. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][33])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
RS2. Spec_GWP_E 	=	Double.parseDouble(data_m[2][34])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
RS2. Spec_AP_E 	=	Double.parseDouble( data_m[3][34]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
RS2. Spec_EP_E 	=	Double.parseDouble( data_m[4][34])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
RS2. Spec_POCP_E 	=	Double.parseDouble(data_m[5][34])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
RS2.H = 7;
RS2.F[0]=Double.parseDouble(data_m1[12][31]);
RS2.F[1]=Double.parseDouble(data_m1[13][31]);
RS2.F[2]=Double.parseDouble(data_m1[14][31]);
RS2.F[3]=Double.parseDouble(data_m1[15][31]);
RS2.F[4]=Double.parseDouble(data_m1[16][31]);
RS2.F[5]=Double.parseDouble(data_m1[17][31]);
RS2.F[6]=Double.parseDouble(data_m1[18][31]);
RS2.C[0]=Double.parseDouble(data_m2[12][31]);
RS2.C[1]=Double.parseDouble(data_m2[13][31]);
RS2.C[2]=Double.parseDouble(data_m2[14][31]);
RS2.C[3]=Double.parseDouble(data_m2[15][31]);
RS2.C[4]=Double.parseDouble(data_m2[16][31]);
RS2.C[5]=Double.parseDouble(data_m2[17][31]);
RS2.C[6]=Double.parseDouble(data_m2[18][31]);
RS2.M[0]=Double.parseDouble(data_m3[12][31]);
RS2.M[1]=Double.parseDouble(data_m3[13][31]);
RS2.M[2]=Double.parseDouble(data_m3[14][31]);
RS2.M[3]=Double.parseDouble(data_m3[15][31]);
RS2.M[4]=Double.parseDouble(data_m3[16][31]);
RS2.M[5]=Double.parseDouble(data_m3[17][31]);
RS2.M[6]=Double.parseDouble(data_m3[18][31]);
RS2.CoTL =CoTL;
RS2.run(); //Run the calculation

Parameter_C_System RS3 = new Parameter_C_System();
RS3.	Engine_type = data_m[0][32];					//battery name
RS3.	Number	=	Double.parseDouble(data_m[1][32])	; //Number of items
RS3.	Weight	=	Double.parseDouble(data_m[2][32])	; //Weight of item (ton)
RS3.	Price	=	Double.parseDouble(data_m[3][32])	; //Item price (Euro/ton)
RS3.	Transportation_type = data_m[0][33];							//transportation type
RS3.	Transportation_distance	=	Double.parseDouble(data_m[1][33])	; //Distance of item transportation (km)
RS3.	Transportation_fee	=	Double.parseDouble(data_m[2][33])	; //Transportation fee per km (Euro/km)
RS3.	Transportation_SFOC	=	Double.parseDouble(data_m[3][33])	; //Transportation fuel consumption (ton/km)
RS3.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][33])	; //Fuel price of transportation (Euro/ton)
RS3.	Electricity_type = data_m[0][34]; //Electricity_type
RS3.	Installation_energy_price	=	Double.parseDouble(data_m[1][34])	; //Energy price (Euro/unit)
RS3. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][33])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
RS3. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][33])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
RS3. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][33])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
RS3. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][33])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
RS3. Spec_GWP_E 	=	Double.parseDouble(data_m[2][34])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
RS3. Spec_AP_E 	=	Double.parseDouble( data_m[3][34]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
RS3. Spec_EP_E 	=	Double.parseDouble( data_m[4][34])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
RS3. Spec_POCP_E 	=	Double.parseDouble(data_m[5][34])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
RS3.H = 7;
RS3.F[0]=Double.parseDouble(data_m1[12][32]);
RS3.F[1]=Double.parseDouble(data_m1[13][32]);
RS3.F[2]=Double.parseDouble(data_m1[14][32]);
RS3.F[3]=Double.parseDouble(data_m1[15][32]);
RS3.F[4]=Double.parseDouble(data_m1[16][32]);
RS3.F[5]=Double.parseDouble(data_m1[17][32]);
RS3.F[6]=Double.parseDouble(data_m1[18][32]);
RS3.C[0]=Double.parseDouble(data_m2[12][32]);
RS3.C[1]=Double.parseDouble(data_m2[13][32]);
RS3.C[2]=Double.parseDouble(data_m2[14][32]);
RS3.C[3]=Double.parseDouble(data_m2[15][32]);
RS3.C[4]=Double.parseDouble(data_m2[16][32]);
RS3.C[5]=Double.parseDouble(data_m2[17][32]);
RS3.C[6]=Double.parseDouble(data_m2[18][32]);
RS3.M[0]=Double.parseDouble(data_m3[12][32]);
RS3.M[1]=Double.parseDouble(data_m3[13][32]);
RS3.M[2]=Double.parseDouble(data_m3[14][32]);
RS3.M[3]=Double.parseDouble(data_m3[15][32]);
RS3.M[4]=Double.parseDouble(data_m3[16][32]);
RS3.M[5]=Double.parseDouble(data_m3[17][32]);
RS3.M[6]=Double.parseDouble(data_m3[18][32]);
RS3.CoTL =CoTL;
RS3.run(); //Run the calculation

//Material: steel, aluminium, composite material...
Parameter_C_Material RM1 = new Parameter_C_Material();
RM1.W1 = 0;
RM1. 	W11 = Double.parseDouble(data_m[	13	][5]);

RM1.	E_price		= Double.parseDouble(data_m[	1	][34]);
RM1.	Spec_GWP_E 	= Double.parseDouble(data_m[	2	][34]);
RM1.	Spec_AP_E 	= Double.parseDouble(data_m[	3	][34]);
RM1.	Spec_EP_E 	= Double.parseDouble(data_m[	4	][34]);
RM1.	Spec_POCP_E = Double.parseDouble(data_m[	5	][34]);

RM1.	P_Cutting	=Double.parseDouble(data_m[	5	][24]);
RM1.	V_Cutting	=Double.parseDouble(data_m[	2	][24]);
RM1.	L_Cutting	=Double.parseDouble(data_m[	1	][24]);
RM1.	SMC_Cutting	=Double.parseDouble(data_m[	3	][24]);
RM1.	MP_Cutting 	= Double.parseDouble(data_m[	4	][24]);
if(Double.parseDouble(data_m[	7	][24])!=0)
{
	RM1.	H_Cutting_Hull	= Double.parseDouble(data_m[	7	][24]);
}
else 
{
	RM1.	r_cutting = Double.parseDouble(data_m[	8	][24]);
}

RM1.P_Bending=Double.parseDouble(data_m[	2	][25]);
RM1.V_Bending=Double.parseDouble(data_m[	1	][25]);
if(Double.parseDouble(data_m[	4	][25])!=0) {
	RM1.H_Bending_Hull= Double.parseDouble(data_m[	4	][25]);
}
else {
	RM1.	r_bending = Double.parseDouble(data_m[	5	][25]);
}

RM1.P_Welding=Double.parseDouble(data_m[	5	][26]);
RM1.V_Welding=Double.parseDouble(data_m[	2	][26]);
RM1.L_Welding=Double.parseDouble(data_m[	1	][26]);
RM1.SMC_Welding=Double.parseDouble(data_m[	3	][26]);
RM1.MP_Welding = Double.parseDouble(data_m[	4	][26]);

//if the hours of welding is inputed just use this hours to determine energy consumption (E=P*H). 
//If it leave as 0, use speed and quantity to find energy consumption (E=P*L/V) 
if(Double.parseDouble(data_m[	7	][26])!=0) 
{
	RM1.H_Welding_Hull= Double.parseDouble(data_m[	7	][26]);
}
else 
{
	RM1.	r_welding = Double.parseDouble(data_m[	8	][26]);
}

RM1.P_Blasting=Double.parseDouble(data_m[	2	][27]);
RM1.V_Blasting=Double.parseDouble(data_m[	1	][27]);
if(Double.parseDouble(data_m[	4	][27])!=0) 
{
	RM1.H_Blasting_Hull= Double.parseDouble(data_m[	4	][27]);
}
else 
{
	RM1.	r_blasting = Double.parseDouble(data_m[	5	][27]);
}

RM1.P_Painting=Double.parseDouble(data_m[	5	][28]);
RM1.V_Painting=Double.parseDouble(data_m[	2	][28]);
RM1.L_Painting=Double.parseDouble(data_m[	1	][28]);
RM1.SMC_Painting=Double.parseDouble(data_m[	3	][28]);
RM1.MP_Painting = Double.parseDouble(data_m[	4	][28]);
if(Double.parseDouble(data_m[	7	][28])!=0) 
{
	RM1.H_Painting_Hull= Double.parseDouble(data_m[	7	][28]);
}
else 
{
	RM1.	r_painting = Double.parseDouble(data_m[	8	][28]);
}

RM1.Transportation_distance=Double.parseDouble(data_m[	1	][29]);
RM1.Transportation_price = Double.parseDouble(data_m[	2	][29]);
RM1.Transportation_SFOC = Double.parseDouble(data_m[	3	][29]);
RM1.Transportation_fuel_price = Double.parseDouble(data_m[	4	][29]);
RM1.Spec_GWP_Trans=Double.parseDouble(data_m[	5	][29]);
RM1.Spec_AP_Trans=Double.parseDouble(data_m[	6	][29]);
RM1.Spec_EP_Trans=Double.parseDouble(data_m[	7	][29]);
RM1.Spec_POCP_Trans=Double.parseDouble(data_m[	8	][29]);

RM1. Labour_fee_cutting =Double.parseDouble(data_m[	6	][24]);
RM1. Labour_fee_bending =Double.parseDouble(data_m[	3	][25]);
RM1. Labour_fee_welding =Double.parseDouble(data_m[	6	][26]);
RM1. Labour_fee_blasting =Double.parseDouble(data_m[	3	][27]);
RM1. Labour_fee_painting =Double.parseDouble(data_m[	6	][28]);

RM1.run(); //Run the calculation


Parameter_C_Material RM2 = new Parameter_C_Material();

RM2.kB=Double.parseDouble(data_m[	1	][36]);
RM2.WB=Double.parseDouble(data_m[	2	][36]);
RM2.E_price=Double.parseDouble(data_m[	1	][34]);
RM2.Spec_GWP_E = Double.parseDouble(data_m[	2	][34]);
RM2.Spec_AP_E = Double.parseDouble(data_m[	3	][34]);
RM2.Spec_EP_E = Double.parseDouble(data_m[	4	][34]);
RM2.Spec_POCP_E = Double.parseDouble(data_m[	5	][34]);

RM2.P_Cutting=Double.parseDouble(data_m[	5	][37]);
RM2.V_Cutting=Double.parseDouble(data_m[	2	][37]);
RM2.L_Cutting=Double.parseDouble(data_m[	1	][37]);
RM2.SMC_Cutting=Double.parseDouble(data_m[	3	][37]);
RM2.MP_Cutting = Double.parseDouble(data_m[	4	][37]);
if(Double.parseDouble(data_m[	7	][37])!=0) 
{
	RM2.H_Cutting_Hull= Double.parseDouble(data_m[	7	][37]);

}
else 
{
	RM2.	r_cutting = Double.parseDouble(data_m[	8	][37]);
}
RM2.P_Welding=Double.parseDouble(data_m[	5	][38]);
RM2.V_Welding=Double.parseDouble(data_m[	2	][38]);
RM2.L_Welding=Double.parseDouble(data_m[	1	][38]);
RM2.SMC_Welding=Double.parseDouble(data_m[	3	][38]);
RM2.MP_Welding = Double.parseDouble(data_m[	4	][38]);
if(Double.parseDouble(data_m[	7	][38])!=0) 
{
	RM2.H_Welding_Hull= Double.parseDouble(data_m[	7	][38]);

}
else 
{
	RM2.	r_welding = Double.parseDouble(data_m[	8	][38]);
}

RM2.P_Painting=Double.parseDouble(data_m[	5	][39]);
RM2.V_Painting=Double.parseDouble(data_m[	2	][39]);
RM2.L_Painting=Double.parseDouble(data_m[	1	][39]);
RM2.SMC_Painting=Double.parseDouble(data_m[	3	][39]);
RM2.MP_Painting = Double.parseDouble(data_m[	4	][39]);
if(Double.parseDouble(data_m[	7	][39])!=0) 
{
	RM2.H_Painting_Hull= Double.parseDouble(data_m[	7	][39]);
}
else 
{
	RM2.	r_painting = Double.parseDouble(data_m[	8	][39]);
}

RM2. Labour_fee_cutting =Double.parseDouble(data_m[	6	][37]);
RM2. Labour_fee_welding =Double.parseDouble(data_m[	6	][38]);
RM2. Labour_fee_painting =Double.parseDouble(data_m[	6	][39]);

RM2.run(); //Run the calculation


//Operation M1
Parameter_O O1 = new Parameter_O();
O1.	Life_span	=	Life_span	;
O1.	Interest	=	Interest	;
O1.	PV 			=	PV	;
O1.	Number 		=	CS1.Number;
O1. Fuel_type 	= 	data_m[0][45];
O1. LO_type  	= 	data_m[0][46];

O1.	Ohour		=	Double.parseDouble(data_m[1][42]); //Operation hours (h)
O1.	Ohour_2		=	Double.parseDouble(data_m[8][42]); //Operation hours (h)
O1.	Ohour_3		=	Double.parseDouble(data_m[15][42]); //Operation hours (h)

O1. Engine_No	=   Double.parseDouble(data_m[2][42]); //Number of operated engines (h)
O1. Engine_No_2	=   Double.parseDouble(data_m[9][42]); //Number of operated engines (h)
O1. Engine_No_3	=   Double.parseDouble(data_m[16][42]); //Number of operated engines (h)

O1.	Eload		=	Double.parseDouble(data_m[3][42])	; //Power (kW) (Engine output)
O1.	Eload_2		=	Double.parseDouble(data_m[10][42])	; //Power (kW) (Engine output)
O1.	Eload_3		=	Double.parseDouble(data_m[17][42])	; //Power (kW) (Engine output)

O1.	SFOC		=	Double.parseDouble(data_m[4][42])	; //Specific fuel oil consumption (g/kWh)
O1.	SFOC_2		=	Double.parseDouble(data_m[11][42])	; //Specific fuel oil consumption (g/kWh)
O1.	SFOC_3		=	Double.parseDouble(data_m[18][42])	; //Specific fuel oil consumption (g/kWh)


O1. C_Factor = Double.parseDouble(data_m[6][45]); 	//carbon content
O1. S_Factor = Double.parseDouble(data_m[7][45]);	//sulfer content
O1. N_Factor = Double.parseDouble(data_m[8][45]);        //nitrogen content
O1.	Fuel_price	=	Double.parseDouble(data_m[1][45])	; //Fuel oil price (Euro/ton)



O1.	SLOC		=	Double.parseDouble(data_m[5][42])	; //Specific lubricating oil consumption (g/kWh)
O1.	SLOC_2		=	Double.parseDouble(data_m[12][42])	; //Specific lubricating oil consumption (g/kWh)
O1.	SLOC_3		=	Double.parseDouble(data_m[19][42])	; //Specific lubricating oil consumption (g/kWh)

O1.	LO_price	=	Double.parseDouble(data_m[1][46])	; //Lub oil price (Euro/ton)
O1.	Transportation_fee	=	Double.parseDouble(data_m[2][47])	; //Transportation fee per km (Euro/km)
O1.	Transportation_SFOC	=	Double.parseDouble(data_m[3][47])	; //Transportation specification fuel consumption (ton/km)
O1.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][47])	; //Fuel price of transportation (Euro/ton)
O1.	Transportation_distance	=	Double.parseDouble(data_m[1][47])	; //Distance of material transportation (km)

O1. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][47])	;
O1. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][47])	;
O1. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][47])	;
O1. Spec_POCP_Trans =	Double.parseDouble(data_m[8][47])	;
O1. Spec_GWP_FO 	=	Double.parseDouble(data_m[2][45])		;
O1. Spec_AP_FO 		=	Double.parseDouble(data_m[3][45])		;
O1. Spec_EP_FO 		=	Double.parseDouble(data_m[4][45])		;
O1. Spec_POCP_FO 	=	Double.parseDouble(data_m[5][45])		;
O1. Spec_GWP_LO 	=	Double.parseDouble(data_m[2][46])		;
O1. Spec_AP_LO 		=	Double.parseDouble(data_m[3][46])		;
O1. Spec_EP_LO 		=	Double.parseDouble(data_m[4][46])		;
O1. Spec_POCP_LO 	=	Double.parseDouble(data_m[5][46])		;
O1.run(); //Run the calculation

//Operation M2
Parameter_O O2 = new Parameter_O();
O2.	Life_span	=	Life_span	;
O2.	Interest	=	Interest	;
O2.	PV 			=	PV	;
O2.	Number 		=	CS1.Number;
O2. Fuel_type 	= 	data_m[0][45];
O2. LO_type  	= 	data_m[0][46];

O2.	Ohour		=	Double.parseDouble(data_m[1][43]); //Operation hours (h)
O2.	Ohour_2		=	Double.parseDouble(data_m[8][43]); //Operation hours (h)
O2.	Ohour_3		=	Double.parseDouble(data_m[15][43]); //Operation hours (h)

O2. Engine_No	=   Double.parseDouble(data_m[2][43]); //Number of operated engines (h)
O2. Engine_No_2	=   Double.parseDouble(data_m[9][43]); //Number of operated engines (h)
O2. Engine_No_3	=   Double.parseDouble(data_m[16][43]); //Number of operated engines (h)

O2.	Eload		=	Double.parseDouble(data_m[3][43])	; //Power (kW) (Engine output)
O2.	Eload_2		=	Double.parseDouble(data_m[10][43])	; //Power (kW) (Engine output)
O2.	Eload_3		=	Double.parseDouble(data_m[17][43])	; //Power (kW) (Engine output)

O2.	SFOC		=	Double.parseDouble(data_m[4][43])	; //Specific fuel oil consumption (g/kWh)
O2.	SFOC_2		=	Double.parseDouble(data_m[11][43])	; //Specific fuel oil consumption (g/kWh)
O2.	SFOC_3		=	Double.parseDouble(data_m[18][43])	; //Specific fuel oil consumption (g/kWh)


O2. C_Factor = Double.parseDouble(data_m[6][45]); 	//carbon content
O2. S_Factor = Double.parseDouble(data_m[7][45]);	//sulfer content
O2. N_Factor = Double.parseDouble(data_m[8][45]);        //nitrogen content
O2.	Fuel_price	=	Double.parseDouble(data_m[1][45])	; //Fuel oil price (Euro/ton)



O2.	SLOC		=	Double.parseDouble(data_m[5][43])	; //Specific lubricating oil consumption (g/kWh)
O2.	SLOC_2		=	Double.parseDouble(data_m[12][43])	; //Specific lubricating oil consumption (g/kWh)
O2.	SLOC_3		=	Double.parseDouble(data_m[19][43])	; //Specific lubricating oil consumption (g/kWh)

O2.	LO_price	=	Double.parseDouble(data_m[1][46])	; //Lub oil price (Euro/ton)
O2.	Transportation_fee	=	Double.parseDouble(data_m[2][47])	; //Transportation fee per km (Euro/km)
O2.	Transportation_SFOC	=	Double.parseDouble(data_m[3][47])	; //Transportation specification fuel consumption (ton/km)
O2.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][47])	; //Fuel price of transportation (Euro/ton)
O2.	Transportation_distance	=	Double.parseDouble(data_m[1][47])	; //Distance of material transportation (km)

//Specific GWP of activity (ton CO2e/ ton fuel used for activity)
//Specific AP of activity (ton SO2e/ ton fuel used for activity)
//Specific EP of activity (ton PO4e/ ton fuel used for activity)
//Specific POCP of activity (ton C2H6e/ ton fuel used for activity)

O2. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][47])	;
O2. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][47])	;
O2. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][47])	;
O2. Spec_POCP_Trans =	Double.parseDouble(data_m[8][47])	;
O2. Spec_GWP_FO 	=	Double.parseDouble(data_m[2][45])		;
O2. Spec_AP_FO 		=	Double.parseDouble(data_m[3][45])		;
O2. Spec_EP_FO 		=	Double.parseDouble(data_m[4][45])		;
O2. Spec_POCP_FO 	=	Double.parseDouble(data_m[5][45])		;
O2. Spec_GWP_LO 	=	Double.parseDouble(data_m[2][46])		;
O2. Spec_AP_LO 		=	Double.parseDouble(data_m[3][46])		;
O2. Spec_EP_LO 		=	Double.parseDouble(data_m[4][46])		;
O2. Spec_POCP_LO 	=	Double.parseDouble(data_m[5][46])		;
O2.run(); //Run the calculation

//Operation M3
Parameter_O O3 = new Parameter_O();
O3.	Life_span	=	Life_span	;
O3.	Interest	=	Interest	;
O3.	PV 			=	PV	;
O3.	Number 		=	Double.parseDouble(data_m[2][44]);
O3. Fuel_type 	= 	data_m[0][45];
O3. LO_type  	= 	data_m[0][46];

if(Double.parseDouble(data_m[	22	][44])==1) {
	O3.	SFOC		=	0	; //Specific fuel oil consumption (g/kWh)
	O3.	SFOC_2		=	0	; //Specific fuel oil consumption (g/kWh)
	O3.	SFOC_3		=	0	; //Specific fuel oil consumption (g/kWh)
	}
else {
	O3.	SFOC		=	Double.parseDouble(data_m[4][44])	; //Specific fuel oil consumption (g/kWh)
	O3.	SFOC_2		=	Double.parseDouble(data_m[11][44])	; //Specific fuel oil consumption (g/kWh)
	O3.	SFOC_3		=	Double.parseDouble(data_m[18][44])	; //Specific fuel oil consumption (g/kWh)
}
//System.out.println("sfoc=" + O3.	SFOC);
O3.	Ohour		=	Double.parseDouble(data_m[1][44]); //Operation hours (h)
O3.	Ohour_2		=	Double.parseDouble(data_m[8][44]); //Operation hours (h)
O3.	Ohour_3		=	Double.parseDouble(data_m[15][44]); //Operation hours (h)

O3. Engine_No	=   Double.parseDouble(data_m[2][44]); //Number of operated engines (h)
O3. Engine_No_2	=   Double.parseDouble(data_m[9][44]); //Number of operated engines (h)
O3. Engine_No_3	=   Double.parseDouble(data_m[16][44]); //Number of operated engines (h)

O3.	Eload		=	Double.parseDouble(data_m[3][44])	; //Power (kW) (Engine output)
O3.	Eload_2		=	Double.parseDouble(data_m[10][44])	; //Power (kW) (Engine output)
O3.	Eload_3		=	Double.parseDouble(data_m[17][44])	; //Power (kW) (Engine output)

O3. C_Factor = Double.parseDouble(data_m[6][45]); 	//carbon content
O3. S_Factor = Double.parseDouble(data_m[7][45]);	//sulfer content
O3. N_Factor = Double.parseDouble(data_m[8][45]);        //nitrogen content
O3.	Fuel_price	=	Double.parseDouble(data_m[1][45])	; //Fuel oil price (Euro/ton)



O3.	SLOC		=	Double.parseDouble(data_m[5][44])	; //Specific lubricating oil consumption (g/kWh)
O3.	SLOC_2		=	Double.parseDouble(data_m[12][44])	; //Specific lubricating oil consumption (g/kWh)
O3.	SLOC_3		=	Double.parseDouble(data_m[19][44])	; //Specific lubricating oil consumption (g/kWh)

O3.	LO_price	=	Double.parseDouble(data_m[1][46])	; //Lub oil price (Euro/ton)
O3.	Transportation_fee	=	Double.parseDouble(data_m[2][47])	; //Transportation fee per km (Euro/km)
O3.	Transportation_SFOC	=	Double.parseDouble(data_m[3][47])	; //Transportation specification fuel consumption (ton/km)
O3.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][47])	; //Fuel price of transportation (Euro/ton)
O3.	Transportation_distance	=	Double.parseDouble(data_m[1][47])	; //Distance of material transportation (km)

//Specific GWP of activity (ton CO2e/ ton fuel used for activity)
//Specific AP of activity (ton SO2e/ ton fuel used for activity)
//Specific EP of activity (ton PO4e/ ton fuel used for activity)
//Specific POCP of activity (ton C2H6e/ ton fuel used for activity)

O3. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][47])	;
O3. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][47])	;
O3. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][47])	;
O3. Spec_POCP_Trans =	Double.parseDouble(data_m[8][47])	;
O3. Spec_GWP_FO 	=	Double.parseDouble(data_m[2][45])		;
O3. Spec_AP_FO 		=	Double.parseDouble(data_m[3][45])		;
O3. Spec_EP_FO 		=	Double.parseDouble(data_m[4][45])		;
O3. Spec_POCP_FO 	=	Double.parseDouble(data_m[5][45])		;
O3. Spec_GWP_LO 	=	Double.parseDouble(data_m[2][46])		;
O3. Spec_AP_LO 		=	Double.parseDouble(data_m[3][46])		;
O3. Spec_EP_LO 		=	Double.parseDouble(data_m[4][46])		;
O3. Spec_POCP_LO 	=	Double.parseDouble(data_m[5][46])		;
O3.Electricity_price = Double.parseDouble(data_m[1][17])		;
O3.run(); //Run the calculation
//Maintenance

Parameter_M M1 = new Parameter_M();
M1.GE_number=Double.parseDouble(data_m[2][42]);
M1.B_number=Double.parseDouble(data_m[2][44]);

M1.Life_span = Life_span;
M1.GE_Power= Double.parseDouble(data_m[4][42]);
M1.GEM_Price= Double.parseDouble(data_m[2][48]);
M1.B_Power= Double.parseDouble(data_m[3][44]);
M1.BM_Price= Double.parseDouble(data_m[3][48]);

M1.	E_Cutting	=Double.parseDouble(data_m[5][6]);
M1.	E_Bending	=Double.parseDouble(data_m[3][7]);
M1.	E_Welding	=Double.parseDouble(data_m[5][8]);
M1. E_Blasting = Double.parseDouble(data_m[3][9]);
M1.	E_Painting	=Double.parseDouble(data_m[5][10]);

M1.	H_Cutting_Ma	=Double.parseDouble(data_m[8][48]);
M1.	H_Bending_Ma	=Double.parseDouble(data_m[9][48]);
M1.	H_Welding_Ma	=Double.parseDouble(data_m[10][48]);
M1.	H_Painting_Ma	=Double.parseDouble(data_m[11][48]);

M1.	H_Cutting_O	=Double.parseDouble(data_m[1][50]);
M1.	H_Bending_O	=Double.parseDouble(data_m[2][50]);
M1.	H_Welding_O	=Double.parseDouble(data_m[3][50]);
M1.	H_Painting_O	=Double.parseDouble(data_m[4][50]);

M1.	H_Cutting_Hull	=Double.parseDouble(data_m[1][49]);
M1.	H_Bending_Hull	=Double.parseDouble(data_m[2][49]);
M1.	H_Welding_Hull	=Double.parseDouble(data_m[4][49]);
M1.	H_Painting_Hull	=Double.parseDouble(data_m[5][49]);
M1.	H_Blasting_Hull	=Double.parseDouble(data_m[3][49]);

M1. Spec_GWP_E = Double.parseDouble(data_m[2][17]);
M1. Spec_AP_E = Double.parseDouble(data_m[3][17]);
M1. Spec_EP_E = Double.parseDouble(data_m[4][17]);
M1. Spec_POCP_E = Double.parseDouble(data_m[5][17]);

M1.E_Working_hours= Double.parseDouble(data_m[1][42]);
M1.Boiler_Working_hours= Double.parseDouble(data_m[1][43]);
M1.GE_Working_hours= Double.parseDouble(data_m[1][42]);
M1.B_Working_hours= Double.parseDouble(data_m[1][44]);


M1.Labour_fee_cutting = Double.parseDouble(data_m[6][6]);
M1.Labour_fee_bending = Double.parseDouble(data_m[4][7]);
M1.Labour_fee_welding = Double.parseDouble(data_m[6][8]);
M1.Labour_fee_blasting = Double.parseDouble(data_m[4][9]);
M1.Labour_fee_coating = Double.parseDouble(data_m[6][10]);

M1.engine_price = Double.parseDouble(data_m[3][12])	;
double Spare_Cost = 0;
System.out.println(Double.parseDouble(data_m[1][52]));
for (int i=1; i<data_length;i++) {
	Spare_Cost+=Double.parseDouble(data_m[i][52]);
	}
M1.Spare_cost =Spare_Cost;
System.out.println(M1.Spare_cost);

M1.run();
//Scrapping
Parameter_S S1 = new Parameter_S();
if(Double.parseDouble(data_m[12][5])!=0) {
S1.Hull_Weight = Double.parseDouble(data_m[12][5]);}
else {
	S1.Hull_Weight = CM1.W1;
}
S1.	Machinery_Number	=	CS1.Number	; //Number of item for scrapping
S1.	Machinery_Weight	=	CS1.Weight	; //Weight of item (ton)
S1.S_Price_M	=	Double.parseDouble(data_m[6][54])	; //Price of machinery scrapping (Euro/set)
S1.S_Price_H	= 	Double.parseDouble(data_m[1][54]);//Price of hullscrapping (Euro/set)
S1.	Transportation_distance	=	Double.parseDouble(data_m[1][55])	; //Distance of material transportation (km)
S1.	Transportation_fee	=	Double.parseDouble(data_m[2][55])	; //Transportation fee per km (Euro/km)
S1.	Transportation_SFOC	=	Double.parseDouble(data_m[3][55]); //Transportation specification fuel consumption (ton/km)
S1.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][55])	; //Fuel price of transportation (Euro/ton)
S1.	Electricity_price	=	Double.parseDouble(data_m[1][56])	;  //Fuel oil price (Euro/ton) natural gas
S1.Cutting_power=Double.parseDouble(data_m[5][6]);
S1.Cutting_hours = Double.parseDouble(data_m[1][57])+Double.parseDouble(data_m[1][58])+Double.parseDouble(data_m[1][59]);
S1.Cleaning_power=Double.parseDouble(data_m[5][10]);
S1.Cleaning_hours = Double.parseDouble(data_m[2][57])+Double.parseDouble(data_m[2][58])+Double.parseDouble(data_m[2][59]);

S1.	Life_span	=	Life_span	;
S1.	Interest	=	Interest	;
S1.	PV 	=PV;
//Specific GWP of activity (ton CO2e/ ton fuel used for activity)
//Specific AP of activity (ton SO2e/ ton fuel used for activity)
//Specific EP of activity (ton PO4e/ ton fuel used for activity)
//Specific POCP of activity (ton C2H6e/ ton fuel used for activity)

S1. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][55])	;
S1. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][55])	;
S1. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][55])	;
S1. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][55])	;
S1. Spec_GWP_E 	=	Double.parseDouble(data_m[2][56])	;
S1. Spec_AP_E 	=	Double.parseDouble(data_m[3][56])	;
S1. Spec_EP_E 	=	Double.parseDouble(data_m[4][56])	;
S1. Spec_POCP_E =	Double.parseDouble(data_m[5][56])	;
S1.run(); //Run the calculation


Parameter_S S2 = new Parameter_S();
S2.	Machinery_Number	=	CS1.Number	; //Number of item for scrapping
S2.	Machinery_Weight	=	CS1.Weight	; //Weight of item (ton)
S2.S_Price_M	=	Double.parseDouble(data_m[8][54])	; //Price of scrapping (Euro/ton)
S2.	Transportation_distance	=	Double.parseDouble(data_m[1][55])	; //Distance of material transportation (km)
S2.	Transportation_fee	=	Double.parseDouble(data_m[2][55])	; //Transportation price per km (Euro/km)
S2.	Transportation_SFOC	=	Double.parseDouble(data_m[3][55]); //Transportation specification fuel consumption (ton/km)
S2.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][55])	; //Fuel price of transportation (Euro/ton)

S2.	Life_span	=	0	;
S2.	Interest	=	Interest	;
S2.	PV 	=PV;
//Specific GWP of activity (ton CO2e/ ton fuel used for activity)
//Specific AP of activity (ton SO2e/ ton fuel used for activity)
//Specific EP of activity (ton PO4e/ ton fuel used for activity)
//Specific POCP of activity (ton C2H6e/ ton fuel used for activity)

S2. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][55])	;
S2. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][55])	;
S2. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][55])	;
S2. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][55])	;
S2. Spec_GWP_E 	=	Double.parseDouble(data_m[2][56])	;
S2. Spec_AP_E 	=	Double.parseDouble(data_m[3][56])	;
S2. Spec_EP_E 	=	Double.parseDouble(data_m[4][56])	;
S2. Spec_POCP_E =	Double.parseDouble(data_m[5][56])	;
S2.run(); //Run the calculation

//Bulk Results
double Total_cost = CM1.Cost_C_Material+CM2.Cost_C_Material+CS1.Cost_C_System + CS2.Cost_C_System+
					RM1.Cost_C_Material+RM2.Cost_C_Material+RS1.Cost_C_System + RS2.Cost_C_System+
					O1.Cost_O+O2.Cost_O+O3.Cost_O+M1.Cost_M+S1.Cost_S+S2.Cost_S; //Total cost (Euro)  	CM1.Cost_C_Material+
double Total_GWP = 	CM1.GWP+CM2.GWP+CS1.GWP+CS2.GWP+
					RM1.GWP+RM2.GWP+RS1.GWP+RS2.GWP+	
					M1.GWP+O1.GWP+O2.GWP+O3.GWP+S1.GWP+S2.GWP; //Total GWP    							CM1.GWP+
double Total_AP = 	CM1.AP+CM2.AP+CS1.AP+CS2.AP+
					RM1.AP+RM2.AP+RS1.AP+RS2.AP+
					M1.AP+O1.AP+O2.AP+O3.AP+S1.AP+S2.AP; //Total AP									CM1.AP+
double Total_EP = 	CM1.EP+CM2.EP+CS1.EP+CS2.EP+
					RM1.EP+RM2.EP+RS1.EP+RS2.EP+
					M1.EP+O1.EP+O2.EP+O3.EP+S1.EP+S2.EP; //Total EP									CM1.EP+
double Total_POCP = CM1.POCP+CM2.POCP+CS1.POCP+CS2.POCP+
					RM1.POCP+RM2.POCP+RS1.POCP+RS2.POCP+
					M1.POCP+O1.POCP+O2.POCP+O3.POCP+S1.POCP+S2.POCP; //Total POCP							CM1.POCP+
Total_RA = CS1.RA+ CS2.RA+RS1.RA+RS2.RA; //Total RA
Total_CRA = Total_RA/1000*CoTL;

//Results breakdown
/*design*/		double sum0 = 0;
/*C_H*/			double sum1 = CM1.Cost_C_Material; 
/*C_M*/			double sum2 = CS1.Cost_C_System+CS2.Cost_C_System ; //+ CM1.Cost_C_Material
/*C_A*/			double sum3 = CM2.Cost_C_Material;
				double sumC = sum1+sum2+sum3;
				/*R_H*/			double sum1r = RM1.Cost_C_Material; 
				/*R_M*/			double sum2r = RS1.Cost_C_System+RS2.Cost_C_System ; //+ CM1.Cost_C_Material
				/*R_A*/			double sum3r = RM2.Cost_C_Material;
								double sumR = sum1r+sum2r+sum3r;
/*O*/			double sum4 = O1.Cost_O+O2.Cost_O+O3.Cost_O;
/*M*/			double sum5 = M1.Cost_M;
/*S*/			double sum6 = S1.Cost_S+S2.Cost_S;
/*Sum*/			sum = Total_cost;
/*design*/		double GWP0 = 0;				
/*C_H*/			double GWP1 = CM1.GWP;
/*C_M*/			double GWP2 = CS1.GWP + CS2.GWP; //+ CM1.Cost_C_Material
/*C_A*/			double GWP3 = CM2.GWP;
				double GWP_C=GWP1+GWP2+GWP3;
				/*R_H*/			double GWP1r = RM1.GWP;
				/*R_M*/			double GWP2r = RS1.GWP + RS2.GWP; //+ CM1.Cost_C_Material
				/*R_A*/			double GWP3r = RM2.GWP;
								double GWP_R=GWP1r+GWP2r+GWP3r;
/*O*/			double GWP4 = O1.GWP+O2.GWP+O3.GWP;
/*M*/			double GWP5 = M1.GWP;
/*S*/			double GWP6 = S1.GWP+S2.GWP;
/*Sum*/			GWP = Total_GWP;
//P_GWP = 24; //Euro per ton
/*design*/		double AP0 = 0;				
/*C_H*/			double AP1 = CM1.AP;
/*C_M*/			double AP2 = CS1.AP +CS2.AP; //+ CM1.Cost_C_Material
/*C_A*/			double AP3 = CM2.AP;
				double AP_C = AP1+AP2+AP3;
				/*R_H*/			double AP1r = RM1.AP;
				/*R_M*/			double AP2r = RS1.AP +RS2.AP; //+ CM1.Cost_C_Material
				/*R_A*/			double AP3r = RM2.AP;
								double AP_R = AP1r+AP2r+AP3r;
/*O*/			double AP4 = O1.AP+O2.AP+O3.AP;
/*M*/			double AP5 = M1.AP;
/*S*/			double AP6 = S1.AP +S2.AP;
/*AP*/			AP = Total_AP;
//P_AP = 7788;
/*design*/		double EP0 = 0;				
/*C_H*/			double EP1 = CM1.EP;
/*C_M*/			double EP2 = CS1.EP +CS2.EP; //+ CM1.Cost_C_Material
/*C_A*/			double EP3 = CM2.EP;
				double EP_C = EP1+EP2+EP3;
				/*R_H*/			double EP1r = RM1.EP;
				/*R_M*/			double EP2r = RS1.EP +RS2.EP; //+ CM1.Cost_C_Material
				/*R_A*/			double EP3r = RM2.EP;
								double EP_R = EP1r+EP2r+EP3r;
/*O*/			double EP4 = O1.EP+O2.EP+O3.EP;
/*M*/			double EP5 = M1.EP;
/*S*/			double EP6 = S1.EP +S2.EP;
/*EP*/			EP = Total_EP;
P_EP = 0;
/*design*/		double POCP0 = 0;				
/*C_H*/			double POCP1 = CM1.POCP;
/*C_M*/			double POCP2 = CS1.POCP +CS2.POCP; //+ CM1.Cost_C_Material
/*C_A*/			double POCP3 = CM2.POCP;
				double POCP_C=POCP1+POCP2+POCP3;
				/*R_H*/			double POCP1r = RM1.POCP;
				/*R_M*/			double POCP2r = RS1.POCP +RS2.POCP; //+ CM1.Cost_C_Material
				/*R_A*/			double POCP3r = RM2.POCP;
								double POCP_R=POCP1r+POCP2r+POCP3r;
/*O*/			double POCP4 = O1.POCP+O2.POCP+O3.POCP;
/*M*/			double POCP5 = M1.POCP;
/*S*/			double POCP6 = S1.POCP +S2.POCP;
/*POCP*/		POCP = Total_POCP;
P_POCP = 0;
/*design*/		double RA0 = 0;				
/*C_H*/			double RA1 = 0;
/*C_M*/			double RA2 = 0; //+ CM1.Cost_C_Material
/*C_A*/			double RA3 = 0;
/*O*/			double RA4 = CS1.RA +CS2.RA;
/*O_r*/			double RA4r = RS1.RA+RS2.RA;	
/*M*/			double RA5 = 0;
/*S*/			double RA6 = 0;
///*RA*/			RA = Total_RA;

double Total_LCTC = Total_cost+Total_GWP*P_GWP+Total_AP*P_AP+Total_EP*P_EP+Total_POCP*P_POCP+Total_CRA;

q.add(String.valueOf(sum0));
q.add(String.valueOf(sum1));
q.add(String.valueOf(sum2));
q.add(String.valueOf(sum3));
q.add(String.valueOf(sum4));
q.add(String.valueOf(sum5));
q.add(String.valueOf(sum6));
q.add(String.valueOf(sum));


textReport.setText( 	"ShipLCA Report"+"\n"+
        "Case Name: "+ project_name +"\n"+
        "Date: " +dateFormat.format(date) +"\n"+"\n"+
        
        PV_state+"\n"+
        SL_state+"\n"+"\n"+
        
        "Life cycle cost (Euro):"+"\n"+
        "	Construction (Structure): 	" + formatter.format(sum1+sum3) +"\n"+
        "	Construction (Machinery): 	" + formatter.format(sum2) +"\n"+
        "	Retrofitting (Structure): 	" + formatter.format(sum1r+sum3r) +"\n"+
        "	Retrofitting (Machinery): 	" + formatter.format(sum2r) +"\n"+
        "	Operation: 		" + formatter.format(sum4) + "\n"+
        "	Maintenance: 	" + formatter.format(sum5) + "\n"+
        "	Scrapping: 		" + formatter.format(sum6) + "\n"+
        "	Total cost: 		" + formatter.format(sum) + "\n"+"\n"+
        
        "Global Warming Potential (GWP) (ton CO2e):"+"\n"+
        "	Construction: 	" + formatter.format(GWP_C) +"\n"+
        "	Retrofitting: 		" + formatter.format(GWP_R) +"\n"+
        "	Operation: 		" + formatter.format(GWP4) +"\n"+
        "	Maintenance: 	" +formatter.format(GWP5) +"\n"+
        "	Scrapping: 		" + formatter.format(GWP6) +"\n"+
        "	Total GWP: 		" + formatter.format(Total_GWP) +"\n"+"\n"+
        
        "Acidification Potential (AP) (ton SO2e):"+"\n"+
        "	Construction: 	" + formatter.format(AP_C)+"\n"+
        "	Retrofitting:		" + formatter.format(AP_R)+"\n"+
        "	Operation: 		" + formatter.format(AP4) +"\n"+
        "	Maintenance: 	" +formatter.format(AP5) +"\n"+
        "	Scrapping: 		" + formatter.format(AP6) +"\n"+
        "	Total AP: 		" + formatter.format(Total_AP) +"\n"+"\n"+
        
        "Eutrophication Potential(EP) (ton PO4e):"+"\n"+
        "	Construction: 	" + formatter.format(EP_C) +"\n"+
        "	Retrofitting:		" + formatter.format(EP_R) +"\n"+
        "	Operation:		" + formatter.format(EP4) +"\n"+
        "	Maintenance: 	" +formatter.format(EP5) +"\n"+
        "	Scrapping: 		" + formatter.format(EP6) +"\n"+
        "	Total EP: 		" + formatter.format(Total_EP) +"\n"+"\n"+
        
        
        "Photochemical Ozone Creation Potential (POCP) (ton C2H6e):"+"\n"+
        "	Construction: 	" + formatter.format(POCP_C) +"\n"+
        "	Retrofitting:		" + formatter.format(POCP_R) +"\n"+
        "	Operation: 		" + formatter.format(POCP4 )+"\n"+
        "	Maintenance: 	" +formatter.format(POCP5) +"\n"+
        "	Scrapping: 		" + formatter.format(POCP6) +"\n"+
        "	Total POCP: 		" + formatter.format(Total_POCP) +"\n"+"\n"+
        
        "Risk Priority Number (RPN) of Machinery  (RPN):"+"\n"+
        "	Construction: 	" + formatter.format(RA2) +"\n"+
        "	Retrofitting:		" + formatter.format(RA2) +"\n"+
        "	Operation: 		" + formatter.format(RA4+RA4r) +"\n"+
        "	Maintenance: 	" +formatter.format(RA5) +"\n"+
        "	Scrapping: 		" + formatter.format(RA6) +"\n"+
        "	Total RPN 		" + formatter.format(Total_RA) +"\n"+"\n"+
        
        
        "Total results"+"\n"+
        "	Life cycle cost (Euro): 	"+ formatter.format(sum) +"\n"+
        "	GWP 	(ton CO2e): 	"+ formatter.format(Total_GWP) +"\n"+
        "	GWP 	(Euro): 	"+ formatter.format(Total_GWP*P_GWP) +"\n"+
        "	AP 	(ton SO2e): 	" + formatter.format(Total_AP) +"\n"+
        "	AP 	(Euro): 	" + formatter.format(Total_AP*P_AP) +"\n"+
        "	EP 	(ton PO4e): 	" + formatter.format(Total_EP) +"\n"+
        "	EP 	(Euro): 	" + formatter.format(Total_EP*P_EP) +"\n"+
        "	POCP 	(ton C2H6e): " + formatter.format(Total_POCP) +"\n"+
        "	POCP 	(Euro): 	" + formatter.format(Total_POCP*P_POCP) +"\n"+
        "	RPN:  		" + formatter.format(Total_RA) +"\n"+
        "	Risk Cost 	(Euro): 	" + formatter.format(Total_CRA) +"\n"+
        "	Life cycle total cost (Euro): 	"+ formatter.format(Total_LCTC) +"\n"+"\n"+
        "*Note: now there is no credit regulation for EP and POCP so these results are zero."+"\n"
);

//add XY charts
DefaultCategoryDataset dataset1 =
        new DefaultCategoryDataset( );
dataset1.addValue( sum1+sum2+sum3 , "Construction" , "" );
dataset1.addValue( sumR , "Retrofitting" , "" );
dataset1.addValue( sum4 , "Operation" , "" );
dataset1.addValue( sum5 , "Maintenance" , "");
dataset1.addValue( sum6 , "Scrapping" , "");
dataset1.addValue( sum , "Total" , "");
// Generate the graph
JFreeChart chart1 = ChartFactory.createBarChart(
        "Life Cycle Cost - " + project_name, // Title
        "Life Stages", // x-axis Label
        "Cost (Euro)", // y-axis Label
        dataset1, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP1 = new ChartPanel(chart1);
panel_chart1 = new JPanel();
panel_chart1.add(CP1);

DefaultCategoryDataset dataset2 =
        new DefaultCategoryDataset( );
dataset2.addValue( GWP_C , "Construction" , "" );
dataset2.addValue( GWP_R , "Retrofitting" , "" );

dataset2.addValue( GWP4 , "Operation" , "" );
dataset2.addValue( GWP5 , "Maintenance" , "" );
dataset2.addValue( GWP6 , "Scrapping" , "" );
dataset2.addValue( Total_GWP , "Total" , "" );
// Generate the graph
JFreeChart chart2 = ChartFactory.createBarChart(
        "Global Warming Potential - " + project_name, // Title
        "Life Stages", // x-axis Label
        "GWP (ton CO2e)", // y-axis Label
        dataset2, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP2 = new ChartPanel(chart2);
panel_chart2 = new JPanel();
panel_chart2.add(CP2);

DefaultCategoryDataset dataset3 =
        new DefaultCategoryDataset( );
dataset3.addValue( AP_C , "Construction" , "" );
dataset3.addValue( AP_R , "Retrofitting" , "" );

dataset3.addValue( AP4 , "Operation" , "" );
dataset3.addValue( AP5 , "Maintenance" , "" );
dataset3.addValue( AP6 , "Scrapping" , "" );
dataset3.addValue( Total_AP , "Total" , "" );
// Generate the graph
JFreeChart chart3 = ChartFactory.createBarChart(
        "Acidification Potential - " + project_name, // Title
        "Life Stages", // x-axis Label
        "AP (ton SO2e)", // y-axis Label
        dataset3, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP3 = new ChartPanel(chart3);
panel_chart3 = new JPanel();
panel_chart3.add(CP3);

DefaultCategoryDataset dataset4 =
        new DefaultCategoryDataset( );
dataset4.addValue( EP_C , "Construction" , "" );
dataset4.addValue( EP_R , "Retrofitting" , "" );

dataset4.addValue( EP4 , "Operation" , "" );
dataset4.addValue( EP5 , "Maintenance" , "" );
dataset4.addValue( EP6 , "Scrapping" , "" );
dataset4.addValue( Total_EP , "Total" , "" );
// Generate the graph
JFreeChart chart4 = ChartFactory.createBarChart(
        "Eutrophication Potential - " + project_name, // Title
        "Life Stages", // x-axis Label
        "EP (ton PO4e)", // y-axis Label
        dataset4, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP4 = new ChartPanel(chart4);
panel_chart4 = new JPanel();
panel_chart4.add(CP4);

DefaultCategoryDataset dataset5 =
        new DefaultCategoryDataset( );
dataset5.addValue( POCP_C , "Construction" , "" );
dataset5.addValue( POCP_R , "Retrofitting" , "" );

dataset5.addValue( POCP4 , "Operation" , "" );
dataset5.addValue( POCP5 , "Maintenance" , "" );
dataset5.addValue( POCP6 , "Scrapping" , "" );
dataset5.addValue( Total_POCP , "Total" , "" );
// Generate the graph
JFreeChart chart5 = ChartFactory.createBarChart(
        "Photochemical Ozone Creation Potential - " + project_name, // Title
        "Life Stages", // x-axis Label
        "POCP (ton C2H6e)", // y-axis Label
        dataset5, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP5 = new ChartPanel(chart5);
panel_chart5 = new JPanel();
panel_chart5.add(CP5);

DefaultCategoryDataset dataset6 =
        new DefaultCategoryDataset( );
dataset6.addValue( RA2 , "Construction" , "" );
dataset6.addValue( RA4r , "Retrofitting" , "" );

dataset6.addValue( RA4 , "Operation" , "" );
dataset6.addValue( RA5 , "Maintenance" , "" );
dataset6.addValue( RA6 , "Scrapping" , "" );
dataset6.addValue( Total_RA , "Total" , "" );
// Generate the graph
JFreeChart chart6 = ChartFactory.createBarChart(
        "Risk Assessment- RPN - " + project_name, // Title
        "Life Stages", // x-axis Label
        "Risk Priority Number", // y-axis Label
        dataset6, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP6 = new ChartPanel(chart6);
panel_chart6 = new JPanel();
panel_chart6.add(CP6);

DefaultCategoryDataset dataset7 =
        new DefaultCategoryDataset( );
dataset7.addValue( sum1+sum3+sum2+(GWP1+GWP3+GWP2)*P_GWP+(AP1+AP3+AP2)*P_AP+(EP1+EP3+EP2)*P_EP+(POCP1+POCP3+POCP2)*P_POCP+RA2/1000*CoTL, "Construction" , "" );        //construction total cost
dataset7.addValue( sumR+GWP_R*P_GWP+AP_R*P_AP+EP_R*P_EP+POCP_R*P_POCP+RA4r/1000*CoTL , "Retrofitting" , "" );        	//Retrofitting total cost

dataset7.addValue( sum4+GWP4*P_GWP+AP4*P_AP+EP4*P_EP+POCP4*P_POCP+RA4/1000*CoTL , "Operation" , "" );        	//operation total cost
dataset7.addValue( sum5+GWP5*P_GWP+AP5*P_AP+EP5*P_EP+POCP5*P_POCP+RA5/1000*CoTL, "Maintenance", "");																		//maintenance total cost
dataset7.addValue( sum6+GWP6*P_GWP+AP6*P_AP+EP6*P_EP+POCP6*P_POCP+RA6/1000*CoTL , "Scrapping" , "" ); 			//scrapping total cost
dataset7.addValue( Total_LCTC , "Total" , "" );     				//life cycle total cost
// Generate the graph
JFreeChart chart7 = ChartFactory.createBarChart(
        "Total Life Cycle Cost - " + project_name, // Title
        "Life Stages", // x-axis Label
        "Cost (Euro)", // y-axis Label
        dataset7, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP7 = new ChartPanel(chart7);
panel_chart7 = new JPanel();

panel_chart7.add(CP7);

Frame = new JFrame("Results plots");

tp_plot1= new JTabbedPane();
tp_plot1.addTab("Cost",ii,panel_chart1,"Cost");
tp_plot1.addTab("GWP",ii,panel_chart2,"GWP");
tp_plot1.addTab("AP",ii,panel_chart3,"AP");
tp_plot1.addTab("EP",ii,panel_chart4,"EP");
tp_plot1.addTab("POCP",ii,panel_chart5,"POCP");
tp_plot1.addTab("RPN",ii,panel_chart6,"RPN");
tp_plot1.addTab("Total cost",ii,panel_chart7,"Total cost");
JPanel panel_plot1 = new JPanel();
panel_plot1.setName("Results plots");
tp_plot1.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
tp_plot1.setTabPlacement(JTabbedPane.TOP);
panel_plot1.add(tp_plot1);

Frame= new JFrame("Results plots");
Frame.add(panel_plot1);
Frame.setExtendedState(JFrame.NORMAL);
Frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
Frame.setResizable(false);
Frame.pack();
Frame.setVisible(true);

F_db_name.dispose();



                            


                            



                                 

FileOutputStream fout;
fout = new FileOutputStream(cwd+"/reports/"+TF_db_name.getText()+".xls");
//Total spreadsheet
@SuppressWarnings("resource")
HSSFWorkbook wb1 = new HSSFWorkbook();
HSSFSheet sheet1 = wb1.createSheet("Total");
HSSFRow row0 = sheet1.createRow(0);
HSSFRow row1 = sheet1.createRow(1);
HSSFRow row2 = sheet1.createRow(2);
HSSFCell item0 = row0.createCell(0);
HSSFCell item1 = row0.createCell(1);
HSSFCell item2 = row0.createCell(2);
HSSFCell item3 = row0.createCell(3);
HSSFCell item4 = row0.createCell(4);
HSSFCell item5 = row0.createCell(5);
HSSFCell num0 = row1.createCell(0);
HSSFCell num1 = row1.createCell(1);
HSSFCell num2 = row1.createCell(2);
HSSFCell num3 = row1.createCell(3);
HSSFCell num4 = row1.createCell(4);
HSSFCell num5 = row1.createCell(5);
HSSFCell unit0 = row2.createCell(0);
HSSFCell unit1 = row2.createCell(1);
HSSFCell unit2 = row2.createCell(2);
HSSFCell unit3 = row2.createCell(3);
HSSFCell unit4 = row2.createCell(4);
HSSFCell unit5 = row2.createCell(5);

num0.setCellType(CellType.STRING);
num1.setCellType(CellType.NUMERIC);
num2.setCellType(CellType.NUMERIC);
num3.setCellType(CellType.NUMERIC);
num4.setCellType(CellType.NUMERIC);
num5.setCellType(CellType.NUMERIC);
num0.setCellValue("Total life cycle cost");
num1.setCellValue(sum1+sum2+sum3+(GWP1+GWP3+GWP2)*P_GWP+(AP1+AP3+AP2)*P_AP+(EP1+EP3+EP2)*P_EP+(POCP1+POCP3+POCP2)*P_POCP+RA2/1000*CoTL);
num2.setCellValue(sum4+GWP4*P_GWP+AP4*P_AP+EP4*P_EP+POCP4*P_POCP+RA4/1000*CoTL);
num3.setCellValue(sum5+GWP5*P_GWP+AP5*P_AP+EP5*P_EP+POCP5*P_POCP+RA5/1000*CoTL);
num4.setCellValue(sum6+GWP6*P_GWP+AP6*P_AP+EP6*P_EP+POCP6*P_POCP+RA6/1000*CoTL);
num5.setCellValue(sum+Total_GWP*P_GWP+Total_AP*P_AP+Total_EP*P_EP+Total_POCP*P_POCP+Total_RA/1000*CoTL);
item0.setCellType(CellType.STRING);
item1.setCellType(CellType.STRING);
item2.setCellType(CellType.STRING);
item3.setCellType(CellType.STRING);
item4.setCellType(CellType.STRING);
item5.setCellType(CellType.STRING);
item0.setCellValue("Life stages");
item1.setCellValue("Construction");
item2.setCellValue("Operation");
item3.setCellValue("Maintenance");
item4.setCellValue("Scrapping");
item5.setCellValue("Total cost");
unit0.setCellType(CellType.STRING);
unit1.setCellType(CellType.STRING);
unit2.setCellType(CellType.STRING);
unit3.setCellType(CellType.STRING);
unit4.setCellType(CellType.STRING);
unit5.setCellType(CellType.STRING);
unit0.setCellValue("Unit");
unit1.setCellValue("Euro");
unit2.setCellValue("Euro");
unit3.setCellValue("Euro");
unit4.setCellValue("Euro");
unit5.setCellValue("Euro");
//Detail report spreadsheet
HSSFSheet sheet2 = wb1.createSheet("Report");
HSSFRow row0_1 = sheet2.createRow(0);
HSSFRow row1_1 = sheet2.createRow(1);
HSSFRow row2_1 = sheet2.createRow(2);

HSSFCell item0_1 = row0_1.createCell(0);
HSSFCell item1_1 = row0_1.createCell(1);
HSSFCell item2_1 = row0_1.createCell(2);
HSSFCell item3_1 = row0_1.createCell(3);
HSSFCell item4_1 = row0_1.createCell(4);
HSSFCell item5_1 = row0_1.createCell(5);
HSSFCell item6_1 = row0_1.createCell(6);
HSSFCell item7_1 = row0_1.createCell(7);
HSSFCell item8_1 = row0_1.createCell(8);
HSSFCell item9_1 = row0_1.createCell(9);
HSSFCell item10_1 = row0_1.createCell(10);
HSSFCell item11_1 = row0_1.createCell(11);
HSSFCell item12_1 = row0_1.createCell(12);
HSSFCell item13_1 = row0_1.createCell(13);
HSSFCell item14_1 = row0_1.createCell(14);
HSSFCell item15_1 = row0_1.createCell(15);
HSSFCell item16_1 = row0_1.createCell(16);
HSSFCell item17_1 = row0_1.createCell(17);
HSSFCell item18_1 = row0_1.createCell(18);
HSSFCell item19_1 = row0_1.createCell(19);
HSSFCell item20_1 = row0_1.createCell(20);
HSSFCell item21_1 = row0_1.createCell(21);
HSSFCell item22_1 = row0_1.createCell(22);
HSSFCell item23_1 = row0_1.createCell(23);
HSSFCell item24_1 = row0_1.createCell(24);
HSSFCell item25_1 = row0_1.createCell(25);
HSSFCell item26_1 = row0_1.createCell(26);
HSSFCell item27_1 = row0_1.createCell(27);
HSSFCell item28_1 = row0_1.createCell(28);
HSSFCell item29_1 = row0_1.createCell(29);
HSSFCell item30_1 = row0_1.createCell(30);
HSSFCell item31_1 = row0_1.createCell(31);
HSSFCell item32_1 = row0_1.createCell(32);
HSSFCell item33_1 = row0_1.createCell(33);
HSSFCell item34_1 = row0_1.createCell(34);
HSSFCell item35_1 = row0_1.createCell(35);
HSSFCell item36_1 = row0_1.createCell(36);
HSSFCell item37_1 = row0_1.createCell(37);
HSSFCell item38_1 = row0_1.createCell(38);
HSSFCell item39_1 = row0_1.createCell(39);
HSSFCell item40_1 = row0_1.createCell(40);
HSSFCell item41_1 = row0_1.createCell(41);

HSSFCell num0_1 = row1_1.createCell(0);
HSSFCell num1_1 = row1_1.createCell(1);
HSSFCell num2_1 = row1_1.createCell(2);
HSSFCell num3_1 = row1_1.createCell(3);
HSSFCell num4_1 = row1_1.createCell(4);
HSSFCell num5_1 = row1_1.createCell(5);
HSSFCell num6_1 = row1_1.createCell(6);
HSSFCell num7_1 = row1_1.createCell(7);
HSSFCell num8_1 = row1_1.createCell(8);
HSSFCell num9_1 = row1_1.createCell(9);
HSSFCell num10_1 = row1_1.createCell(10);
HSSFCell num11_1 = row1_1.createCell(11);
HSSFCell num12_1 = row1_1.createCell(12);
HSSFCell num13_1 = row1_1.createCell(13);
HSSFCell num14_1 = row1_1.createCell(14);
HSSFCell num15_1 = row1_1.createCell(15);
HSSFCell num16_1 = row1_1.createCell(16);
HSSFCell num17_1 = row1_1.createCell(17);
HSSFCell num18_1 = row1_1.createCell(18);
HSSFCell num19_1 = row1_1.createCell(19);
HSSFCell num20_1 = row1_1.createCell(20);
HSSFCell num21_1 = row1_1.createCell(21);
HSSFCell num22_1 = row1_1.createCell(22);
HSSFCell num23_1 = row1_1.createCell(23);
HSSFCell num24_1 = row1_1.createCell(24);
HSSFCell num25_1 = row1_1.createCell(25);
HSSFCell num26_1 = row1_1.createCell(26);
HSSFCell num27_1 = row1_1.createCell(27);
HSSFCell num28_1 = row1_1.createCell(28);
HSSFCell num29_1 = row1_1.createCell(29);
HSSFCell num30_1 = row1_1.createCell(30);
HSSFCell num31_1 = row1_1.createCell(31);
HSSFCell num32_1 = row1_1.createCell(32);
HSSFCell num33_1 = row1_1.createCell(33);
HSSFCell num34_1 = row1_1.createCell(34);
HSSFCell num35_1 = row1_1.createCell(35);
HSSFCell num36_1 = row1_1.createCell(36);
HSSFCell num37_1 = row1_1.createCell(37);
HSSFCell num38_1 = row1_1.createCell(38);
HSSFCell num39_1 = row1_1.createCell(39);
HSSFCell num40_1 = row1_1.createCell(40);
HSSFCell num41_1 = row1_1.createCell(41);

HSSFCell unit0_1 = row2_1.createCell(0);
HSSFCell unit1_1 = row2_1.createCell(1);
HSSFCell unit2_1 = row2_1.createCell(2);
HSSFCell unit3_1 = row2_1.createCell(3);
HSSFCell unit4_1 = row2_1.createCell(4);
HSSFCell unit5_1 = row2_1.createCell(5);
HSSFCell unit6_1 = row2_1.createCell(6);
HSSFCell unit7_1 = row2_1.createCell(7);
HSSFCell unit8_1 = row2_1.createCell(8);
HSSFCell unit9_1 = row2_1.createCell(9);
HSSFCell unit10_1 = row2_1.createCell(10);
HSSFCell unit11_1 = row2_1.createCell(11);
HSSFCell unit12_1 = row2_1.createCell(12);
HSSFCell unit13_1 = row2_1.createCell(13);
HSSFCell unit14_1 = row2_1.createCell(14);
HSSFCell unit15_1 = row2_1.createCell(15);
HSSFCell unit16_1 = row2_1.createCell(16);
HSSFCell unit17_1 = row2_1.createCell(17);
HSSFCell unit18_1 = row2_1.createCell(18);
HSSFCell unit19_1 = row2_1.createCell(19);
HSSFCell unit20_1 = row2_1.createCell(20);
HSSFCell unit21_1 = row2_1.createCell(21);
HSSFCell unit22_1 = row2_1.createCell(22);
HSSFCell unit23_1 = row2_1.createCell(23);
HSSFCell unit24_1 = row2_1.createCell(24);
HSSFCell unit25_1 = row2_1.createCell(25);
HSSFCell unit26_1 = row2_1.createCell(26);
HSSFCell unit27_1 = row2_1.createCell(27);
HSSFCell unit28_1 = row2_1.createCell(28);
HSSFCell unit29_1 = row2_1.createCell(29);
HSSFCell unit30_1 = row2_1.createCell(30);
HSSFCell unit31_1 = row2_1.createCell(31);
HSSFCell unit32_1 = row2_1.createCell(32);
HSSFCell unit33_1 = row2_1.createCell(33);
HSSFCell unit34_1 = row2_1.createCell(34);
HSSFCell unit35_1 = row2_1.createCell(35);
HSSFCell unit36_1 = row2_1.createCell(36);
HSSFCell unit37_1 = row2_1.createCell(37);
HSSFCell unit38_1 = row2_1.createCell(38);
HSSFCell unit39_1 = row2_1.createCell(39);
HSSFCell unit40_1 = row2_1.createCell(40);
HSSFCell unit41_1 = row2_1.createCell(41);


num0_1.setCellType(CellType.STRING);
num0_1.setCellValue("Value");
num1_1.setCellType(CellType.NUMERIC);
num2_1.setCellType(CellType.NUMERIC);
num3_1.setCellType(CellType.NUMERIC);
num4_1.setCellType(CellType.NUMERIC);
num5_1.setCellType(CellType.NUMERIC);
num6_1.setCellType(CellType.NUMERIC);
num7_1.setCellType(CellType.NUMERIC);
num8_1.setCellType(CellType.NUMERIC);
num9_1.setCellType(CellType.NUMERIC);
num10_1.setCellType(CellType.NUMERIC);
num11_1.setCellType(CellType.NUMERIC);
num12_1.setCellType(CellType.NUMERIC);
num13_1.setCellType(CellType.NUMERIC);
num14_1.setCellType(CellType.NUMERIC);
num15_1.setCellType(CellType.NUMERIC);
num16_1.setCellType(CellType.NUMERIC);
num17_1.setCellType(CellType.NUMERIC);
num18_1.setCellType(CellType.NUMERIC);
num19_1.setCellType(CellType.NUMERIC);
num20_1.setCellType(CellType.NUMERIC);
num21_1.setCellType(CellType.NUMERIC);
num22_1.setCellType(CellType.NUMERIC);
num23_1.setCellType(CellType.NUMERIC);
num24_1.setCellType(CellType.NUMERIC);
num25_1.setCellType(CellType.NUMERIC);
num26_1.setCellType(CellType.NUMERIC);
num27_1.setCellType(CellType.NUMERIC);
num28_1.setCellType(CellType.NUMERIC);
num29_1.setCellType(CellType.NUMERIC);
num30_1.setCellType(CellType.NUMERIC);
num31_1.setCellType(CellType.NUMERIC);
num32_1.setCellType(CellType.NUMERIC);
num33_1.setCellType(CellType.NUMERIC);
num34_1.setCellType(CellType.NUMERIC);
num35_1.setCellType(CellType.NUMERIC);
num36_1.setCellType(CellType.NUMERIC);
num37_1.setCellType(CellType.NUMERIC);
num38_1.setCellType(CellType.NUMERIC);
num39_1.setCellType(CellType.NUMERIC);
num40_1.setCellType(CellType.NUMERIC);
num41_1.setCellType(CellType.NUMERIC);

num1_1.setCellValue( sum1+sum2+sum3 );
num2_1.setCellValue( sum4 );
num3_1.setCellValue( sum5 );

num4_1.setCellValue( sum6 );
num5_1.setCellValue( sum );
num6_1.setCellValue( GWP_C );
num7_1.setCellValue( GWP4 );
num8_1.setCellValue( GWP5 );

num9_1.setCellValue( GWP6 );
num10_1.setCellValue( Total_GWP );
num11_1.setCellValue( AP_C );
num12_1.setCellValue( AP4 );
num13_1.setCellValue( AP5 );

num14_1.setCellValue( AP6 );
num15_1.setCellValue( Total_AP );
num16_1.setCellValue( EP_C );
num17_1.setCellValue( EP4 );
num18_1.setCellValue( EP5 );

num19_1.setCellValue( EP6 );
num20_1.setCellValue( Total_EP );
num21_1.setCellValue( POCP_C );
num22_1.setCellValue( POCP4 );
num23_1.setCellValue( POCP5 );

num24_1.setCellValue( POCP6 );
num25_1.setCellValue( Total_POCP );
num26_1.setCellValue( RA2 );
num27_1.setCellValue( RA4 );
num28_1.setCellValue( RA5 );

num29_1.setCellValue( RA6 );
num30_1.setCellValue( Total_RA );
num31_1.setCellValue( Total_cost );
num32_1.setCellValue( Total_GWP );
num33_1.setCellValue( Total_AP );
num34_1.setCellValue( Total_EP );
num35_1.setCellValue( Total_POCP );
num36_1.setCellValue( Total_RA );
num37_1.setCellValue( Total_CRA );
num38_1.setCellValue( Total_GWP*P_GWP );
num39_1.setCellValue( Total_AP*P_AP );
num40_1.setCellValue( Total_EP*P_EP );
num41_1.setCellValue( Total_POCP*P_POCP );

item0_1.setCellType(CellType.STRING);
item1_1.setCellType(CellType.STRING);
item2_1.setCellType(CellType.STRING);
item3_1.setCellType(CellType.STRING);
item4_1.setCellType(CellType.STRING);
item5_1.setCellType(CellType.STRING);
item6_1.setCellType(CellType.STRING);
item7_1.setCellType(CellType.STRING);
item8_1.setCellType(CellType.STRING);
item9_1.setCellType(CellType.STRING);
item10_1.setCellType(CellType.STRING);
item11_1.setCellType(CellType.STRING);
item12_1.setCellType(CellType.STRING);
item13_1.setCellType(CellType.STRING);
item14_1.setCellType(CellType.STRING);
item15_1.setCellType(CellType.STRING);
item16_1.setCellType(CellType.STRING);
item17_1.setCellType(CellType.STRING);
item18_1.setCellType(CellType.STRING);
item19_1.setCellType(CellType.STRING);
item20_1.setCellType(CellType.STRING);
item21_1.setCellType(CellType.STRING);
item22_1.setCellType(CellType.STRING);
item23_1.setCellType(CellType.STRING);
item24_1.setCellType(CellType.STRING);
item25_1.setCellType(CellType.STRING);
item26_1.setCellType(CellType.STRING);
item27_1.setCellType(CellType.STRING);
item28_1.setCellType(CellType.STRING);
item29_1.setCellType(CellType.STRING);
item30_1.setCellType(CellType.STRING);
item31_1.setCellType(CellType.STRING);
item32_1.setCellType(CellType.STRING);
item33_1.setCellType(CellType.STRING);
item34_1.setCellType(CellType.STRING);
item35_1.setCellType(CellType.STRING);
item36_1.setCellType(CellType.STRING);
item37_1.setCellType(CellType.STRING);
item38_1.setCellType(CellType.STRING);
item39_1.setCellType(CellType.STRING);
item40_1.setCellType(CellType.STRING);
item41_1.setCellType(CellType.STRING);

item0_1.setCellValue("	Phases & items	");
item1_1.setCellValue("	Construction cost	");
item2_1.setCellValue("	Operation cost	");
item3_1.setCellValue("	Maintenance cost  	");
item4_1.setCellValue("	Scrapping cost  	");
item5_1.setCellValue("	Total cost  	");

item6_1.setCellValue("	Construction GWP  	");
item7_1.setCellValue("	Operation GWP  	");
item8_1.setCellValue("	Maintenance GWP  	");
item9_1.setCellValue("	Scrapping GWP  	");
item10_1.setCellValue("	Total GWP  	");

item11_1.setCellValue("	Construction AP  	");
item12_1.setCellValue("	Operation AP 	");
item13_1.setCellValue("	Maintenance AP  	");
item14_1.setCellValue("	Scrapping AP  	");
item15_1.setCellValue("	Total AP  	");

item16_1.setCellValue("	Construction EP  	");
item17_1.setCellValue("	Operation EP  	");
item18_1.setCellValue("	Maintenance EP	");
item19_1.setCellValue("	Scrapping EP	");
item20_1.setCellValue("	Total EP	");

item21_1.setCellValue("	Construction POCP	");
item22_1.setCellValue("	Operation POCP	");
item23_1.setCellValue("	Maintenance POCP	");
item24_1.setCellValue("	Scrapping POCP	");
item25_1.setCellValue("	Total POCP	");

item26_1.setCellValue("	Construction RPN	");
item27_1.setCellValue("	Operation RPN	");
item28_1.setCellValue("	Maintenance RPN	");
item29_1.setCellValue("	Scrapping RPN	");
item30_1.setCellValue("	Total RPN	");

item31_1.setCellValue("	Life Cycle Cost 	");
item32_1.setCellValue("	GWP 	");
item33_1.setCellValue("	AP 	");
item34_1.setCellValue("	EP	");
item35_1.setCellValue("	POCP	");
item36_1.setCellValue("	RPN	");

item37_1.setCellValue("	Risk Cost	");
item38_1.setCellValue("	GWP 	");
item39_1.setCellValue("	AP 	");
item40_1.setCellValue("	EP	");
item41_1.setCellValue("	POCP	");

unit0_1.setCellType(CellType.STRING);
unit1_1.setCellType(CellType.STRING);
unit2_1.setCellType(CellType.STRING);
unit3_1.setCellType(CellType.STRING);
unit4_1.setCellType(CellType.STRING);
unit5_1.setCellType(CellType.STRING);
unit6_1.setCellType(CellType.STRING);
unit7_1.setCellType(CellType.STRING);
unit8_1.setCellType(CellType.STRING);
unit9_1.setCellType(CellType.STRING);
unit10_1.setCellType(CellType.STRING);
unit11_1.setCellType(CellType.STRING);
unit12_1.setCellType(CellType.STRING);
unit13_1.setCellType(CellType.STRING);
unit14_1.setCellType(CellType.STRING);
unit15_1.setCellType(CellType.STRING);
unit16_1.setCellType(CellType.STRING);
unit17_1.setCellType(CellType.STRING);
unit18_1.setCellType(CellType.STRING);
unit19_1.setCellType(CellType.STRING);
unit20_1.setCellType(CellType.STRING);
unit21_1.setCellType(CellType.STRING);
unit22_1.setCellType(CellType.STRING);
unit23_1.setCellType(CellType.STRING);
unit24_1.setCellType(CellType.STRING);
unit25_1.setCellType(CellType.STRING);
unit26_1.setCellType(CellType.STRING);
unit27_1.setCellType(CellType.STRING);
unit28_1.setCellType(CellType.STRING);
unit29_1.setCellType(CellType.STRING);
unit30_1.setCellType(CellType.STRING);
unit31_1.setCellType(CellType.STRING);
unit32_1.setCellType(CellType.STRING);
unit33_1.setCellType(CellType.STRING);
unit34_1.setCellType(CellType.STRING);
unit35_1.setCellType(CellType.STRING);
unit35_1.setCellType(CellType.STRING);
unit36_1.setCellType(CellType.STRING);
unit37_1.setCellType(CellType.STRING);
unit38_1.setCellType(CellType.STRING);
unit39_1.setCellType(CellType.STRING);
unit40_1.setCellType(CellType.STRING);
unit41_1.setCellType(CellType.STRING);

unit0_1.setCellValue("Units");
unit1_1.setCellValue("Euro");
unit2_1.setCellValue("Euro");
unit3_1.setCellValue("Euro");
unit4_1.setCellValue("Euro");
unit5_1.setCellValue("Euro");

unit6_1.setCellValue("ton CO2e");
unit7_1.setCellValue("ton CO2e");
unit8_1.setCellValue("ton CO2e");
unit9_1.setCellValue("ton CO2e");
unit10_1.setCellValue("ton CO2e");

unit11_1.setCellValue("ton SO2e");
unit12_1.setCellValue("ton SO2e");
unit13_1.setCellValue("ton SO2e");
unit14_1.setCellValue("ton SO2e");
unit15_1.setCellValue("ton SO2e");

unit16_1.setCellValue("ton PO4e");
unit17_1.setCellValue("ton PO4e");
unit18_1.setCellValue("ton PO4e");
unit19_1.setCellValue("ton PO4e");
unit20_1.setCellValue("ton PO4e");

unit21_1.setCellValue("ton C2H6e");
unit22_1.setCellValue("ton C2H6e");
unit23_1.setCellValue("ton C2H6e");
unit24_1.setCellValue("ton C2H6e");
unit25_1.setCellValue("ton C2H6e");

unit26_1.setCellValue("RPN");
unit27_1.setCellValue("RPN");
unit28_1.setCellValue("RPN");
unit29_1.setCellValue("RPN");
unit30_1.setCellValue("RPN");

unit31_1.setCellValue("Euro");
unit32_1.setCellValue("ton CO2e");
unit33_1.setCellValue("ton SO2e");
unit34_1.setCellValue("ton PO4e");
unit35_1.setCellValue("ton C2H6e");
unit36_1.setCellValue("RPN");
unit37_1.setCellValue("Euro");
unit38_1.setCellValue("Euro");
unit39_1.setCellValue("Euro");
unit40_1.setCellValue("Euro");
unit41_1.setCellValue("Euro");

//Detail report spreadsheet
HSSFSheet sheet3 = wb1.createSheet("Retrofitting");
HSSFRow row0_2 = sheet3.createRow(0);
HSSFRow row1_2 = sheet3.createRow(1);
HSSFRow row2_2 = sheet3.createRow(2);

HSSFCell item0_2 = row0_2.createCell(0);
HSSFCell item1_2 = row0_2.createCell(1);
HSSFCell item2_2 = row0_2.createCell(2);
HSSFCell item3_2 = row0_2.createCell(3);
HSSFCell item4_2 = row0_2.createCell(4);
HSSFCell item5_2 = row0_2.createCell(5);
HSSFCell item6_2 = row0_2.createCell(6);


HSSFCell num0_2 = row1_2.createCell(0);
HSSFCell num1_2 = row1_2.createCell(1);
HSSFCell num2_2 = row1_2.createCell(2);
HSSFCell num3_2 = row1_2.createCell(3);
HSSFCell num4_2 = row1_2.createCell(4);
HSSFCell num5_2 = row1_2.createCell(5);
HSSFCell num6_2 = row1_2.createCell(6);

HSSFCell unit0_2 = row2_2.createCell(0);
HSSFCell unit1_2 = row2_2.createCell(1);
HSSFCell unit2_2 = row2_2.createCell(2);
HSSFCell unit3_2 = row2_2.createCell(3);
HSSFCell unit4_2 = row2_2.createCell(4);
HSSFCell unit5_2 = row2_2.createCell(5);
HSSFCell unit6_2 = row2_2.createCell(6);


num0_2.setCellType(CellType.STRING);
num0_2.setCellValue("Value");
num1_2.setCellType(CellType.NUMERIC);
num2_2.setCellType(CellType.NUMERIC);
num3_2.setCellType(CellType.NUMERIC);
num4_2.setCellType(CellType.NUMERIC);
num5_2.setCellType(CellType.NUMERIC);
num6_2.setCellType(CellType.NUMERIC);


num1_2.setCellValue( sumR );
num2_2.setCellValue( GWP_R );
num3_2.setCellValue( AP_R );
num4_2.setCellValue( EP_R );
num5_2.setCellValue( POCP_R );
num6_2.setCellValue( RA4r );

item0_2.setCellType(CellType.STRING);
item1_2.setCellType(CellType.STRING);
item2_2.setCellType(CellType.STRING);
item3_2.setCellType(CellType.STRING);
item4_2.setCellType(CellType.STRING);
item5_2.setCellType(CellType.STRING);
item6_2.setCellType(CellType.STRING);


item0_2.setCellValue("	Retrofitting	");
item1_2.setCellValue("	Life cycle cost	");
item2_2.setCellValue("	GWP	");
item3_2.setCellValue("	AP  	");
item4_2.setCellValue("	EP  	");
item5_2.setCellValue("	POCP  	");

item6_2.setCellValue("	RPN  	");


unit0_2.setCellType(CellType.STRING);
unit1_2.setCellType(CellType.STRING);
unit2_2.setCellType(CellType.STRING);
unit3_2.setCellType(CellType.STRING);
unit4_2.setCellType(CellType.STRING);
unit5_2.setCellType(CellType.STRING);
unit6_2.setCellType(CellType.STRING);


unit0_2.setCellValue("Units");
unit1_2.setCellValue("Euro");
unit2_2.setCellValue("tonne CO2e");
unit3_2.setCellValue("tonne SO2e");
unit4_2.setCellValue("tonne PO4e.");
unit5_2.setCellValue("tonne C2H6e.");
unit6_2.setCellValue("RPN number");



wb1.write(fout);
F_db_name.dispose();


                            } catch (FileNotFoundException ex) {
                                java.util.logging.Logger.getLogger(Gui16052019.class.getName()).log(Level.SEVERE, null, ex);
                            } catch (IOException ex) {
                                java.util.logging.Logger.getLogger(Gui16052019.class.getName()).log(Level.SEVERE, null, ex);
                            } catch (Exception ex) {
                                java.util.logging.Logger.getLogger(Gui16052019.class.getName()).log(Level.SEVERE, null, ex);
                                JFrame F_warn0 = new JFrame("Error!");
								JPanel P_warn0 = new JPanel();
								P_warn0.setLayout(new BorderLayout());
								JLabel Fd_warn0 = new JLabel("Error: please make sure there are no 'BLANK' cell and run again!");
								F_warn0.setSize(500, 80);
								//Fd_warn0.setText("Database saved in:" + "\n"+ cwd+"/db/" + cb_m[j].getName()+"/"+dateFormat.format(d)+".xls");
								P_warn0.add(Fd_warn0,BorderLayout.CENTER);
								F_warn0.setLocation(400, 170);
								F_warn0.add(P_warn0);
								F_warn0.setVisible(true);
								F_warn0.setResizable(false);
                            }
					
					}
				}
				);
				
				cancel_name.addActionListener(new ActionListener() {
					
					@Override
					public void actionPerformed(ActionEvent e) {
				
                            try { 
                            	
PV =Double.parseDouble(data_m1[1][2]); //0 means not using present value; 1 means using;
Life_span = Double.parseDouble(field0[1].getText()); //Life span of target (year)
Interest= Double.parseDouble(data_m1[2][2]); //Interest rate (100%)
P_GWP = Double.parseDouble(data_m1[3][2]);
P_AP = Double.parseDouble(data_m1[4][2]);
project_name = field0[0].getText();						//project name
SL= Double.parseDouble(field0[3].getText());			//sensitivity level
CoTL = Double.parseDouble(field0[4].getText());		//ship total price
String PV_state, SL_state = null;
//Statement to identify PV/SL
if(PV==0)
{
	PV_state = "Present value is not applied;";
}
else
{
	PV_state = "Present value is applied;";
}

if(SL==0)
{
	data_m=data_m1;
	SL_state="Average value is applied;";
	System.out.println(SL_state);
}


if(SL==1)
{
	data_m=data_m2;
	SL_state="Minimum value is applied;";
	System.out.println(SL_state);
}


if(SL==2)
{
	data_m=data_m3;
	SL_state="Maximum value is applied;";
	System.out.println(SL_state);
}


System.out.println("test"+data_m[1][12]);
//Construction
//System: ME, MG, Boiler...
Parameter_C_System CS1 = new Parameter_C_System();
CS1.	Engine_type	=	data_m[0][12];					//Engine type
CS1.	Number	=	Double.parseDouble(data_m[1][12])	; //Number of items
CS1.	Weight	=	Double.parseDouble(data_m[2][12])	; //Weight of item (ton)
CS1.	Price	=	Double.parseDouble(data_m[3][12])	; //Item price (Euro/ton)
CS1.	Transportation_type = cb_m[15].getName();							//transportation type
CS1.	Transportation_distance	=	Double.parseDouble(data_m[1][15])	; //Distance of item transportation (km)
CS1.	Transportation_fee	=	Double.parseDouble(data_m[2][15])	; //Transportation price per km (Euro/km)
CS1.	Transportation_SFOC	=	Double.parseDouble(data_m[3][15])	; //Transportation fuel consumption (ton/km)
CS1.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][15])	; //Fuel price of transportation (Euro/ton)
CS1.	Electricity_type = cb_m[17].getName(); //Electricity_type
CS1.	Installation_energy_price	=	Double.parseDouble(data_m[1][17])	; //Energy price (Euro/unit)
CS1. 	Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][15])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
CS1. 	Spec_AP_Trans 	=	Double.parseDouble(data_m[6][15])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
CS1. 	Spec_EP_Trans 	=	Double.parseDouble(data_m[7][15])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
CS1. 	Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][15])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
CS1. 	Spec_GWP_E 	=	Double.parseDouble(data_m[2][17])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
CS1. 	Spec_AP_E 	=	Double.parseDouble( data_m[3][17]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
CS1. 	Spec_EP_E 	=	Double.parseDouble( data_m[4][17])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
CS1. 	Spec_POCP_E 	=	Double.parseDouble(data_m[5][17])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
CS1.H = 7;
CS1.F[0]=Double.parseDouble(data_m1[10][12]);
CS1.F[1]=Double.parseDouble(data_m1[11][12]);
CS1.F[2]=Double.parseDouble(data_m1[12][12]);
CS1.F[3]=Double.parseDouble(data_m1[13][12]);
CS1.F[4]=Double.parseDouble(data_m1[14][12]);
CS1.F[5]=Double.parseDouble(data_m1[15][12]);
CS1.F[6]=Double.parseDouble(data_m1[16][12]);
CS1.C[0]=Double.parseDouble(data_m2[10][12]);
CS1.C[1]=Double.parseDouble(data_m2[11][12]);
CS1.C[2]=Double.parseDouble(data_m2[12][12]);
CS1.C[3]=Double.parseDouble(data_m2[13][12]);
CS1.C[4]=Double.parseDouble(data_m2[14][12]);
CS1.C[5]=Double.parseDouble(data_m2[15][12]);
CS1.C[6]=Double.parseDouble(data_m2[16][12]);
CS1.M[0]=Double.parseDouble(data_m3[10][12]);
CS1.M[1]=Double.parseDouble(data_m3[11][12]);
CS1.M[2]=Double.parseDouble(data_m3[12][12]);
CS1.M[3]=Double.parseDouble(data_m3[13][12]);
CS1.M[4]=Double.parseDouble(data_m3[14][12]);
CS1.M[5]=Double.parseDouble(data_m3[15][12]);
CS1.M[6]=Double.parseDouble(data_m3[16][12]);
CS1.CoTL =CoTL;
CS1.run(); //Run the calculation
Parameter_C_System CS2 = new Parameter_C_System();
CS2.	Engine_type = data_m[0][14];					//battery name
CS2.	Number	=	Double.parseDouble(data_m[1][14])	; //Number of items
CS2.	Weight	=	Double.parseDouble(data_m[2][14])	; //Weight of item (ton)
CS2.	Price	=	Double.parseDouble(data_m[3][14])	; //Item price (Euro/ton)
CS2.	Transportation_type = data_m[0][16];							//transportation type
CS2.	Transportation_distance	=	Double.parseDouble(data_m[1][16])	; //Distance of item transportation (km)
CS2.	Transportation_fee	=	Double.parseDouble(data_m[2][16])	; //Transportation fee per km (Euro/km)
CS2.	Transportation_SFOC	=	Double.parseDouble(data_m[3][16])	; //Transportation fuel consumption (ton/km)
CS2.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][16])	; //Fuel price of transportation (Euro/ton)
CS2.	Electricity_type = data_m[0][17]; //Electricity_type
CS2.	Installation_energy_price	=	Double.parseDouble(data_m[1][17])	; //Energy price (Euro/unit)
CS2. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][16])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
CS2. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][16])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
CS2. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][16])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
CS2. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][16])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
CS2. Spec_GWP_E 	=	Double.parseDouble(data_m[2][17])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
CS2. Spec_AP_E 	=	Double.parseDouble( data_m[3][17]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
CS2. Spec_EP_E 	=	Double.parseDouble( data_m[4][17])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
CS2. Spec_POCP_E 	=	Double.parseDouble(data_m[5][17])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
CS2.H = 7;
CS2.F[0]=Double.parseDouble(data_m1[12][14]);
CS2.F[1]=Double.parseDouble(data_m1[13][14]);
CS2.F[2]=Double.parseDouble(data_m1[14][14]);
CS2.F[3]=Double.parseDouble(data_m1[15][14]);
CS2.F[4]=Double.parseDouble(data_m1[16][14]);
CS2.F[5]=Double.parseDouble(data_m1[17][14]);
CS2.F[6]=Double.parseDouble(data_m1[18][14]);
CS2.C[0]=Double.parseDouble(data_m2[12][14]);
CS2.C[1]=Double.parseDouble(data_m2[13][14]);
CS2.C[2]=Double.parseDouble(data_m2[14][14]);
CS2.C[3]=Double.parseDouble(data_m2[15][14]);
CS2.C[4]=Double.parseDouble(data_m2[16][14]);
CS2.C[5]=Double.parseDouble(data_m2[17][14]);
CS2.C[6]=Double.parseDouble(data_m2[18][14]);
CS2.M[0]=Double.parseDouble(data_m3[12][14]);
CS2.M[1]=Double.parseDouble(data_m3[13][14]);
CS2.M[2]=Double.parseDouble(data_m3[14][14]);
CS2.M[3]=Double.parseDouble(data_m3[15][14]);
CS2.M[4]=Double.parseDouble(data_m3[16][14]);
CS2.M[5]=Double.parseDouble(data_m3[17][14]);
CS2.M[6]=Double.parseDouble(data_m3[18][14]);
CS2.CoTL =CoTL;
CS2.run(); //Run the calculation


Parameter_C_System CS3 = new Parameter_C_System();
CS3.	Engine_type = data_m[0][13];					//Genset name
CS3.	Number	=	Double.parseDouble(data_m[1][13])	; //Number of items
CS3.	Weight	=	Double.parseDouble(data_m[2][13])	; //Weight of item (ton)
CS3.	Price	=	Double.parseDouble(data_m[3][13])	; //Item price (Euro/ton)
CS3.	Transportation_type = data_m[0][15];							//transportation type
CS3.	Transportation_distance	=	Double.parseDouble(data_m[1][15])	; //Distance of item transportation (km)
CS3.	Transportation_fee	=	Double.parseDouble(data_m[2][15])	; //Transportation fee per km (Euro/km)
CS3.	Transportation_SFOC	=	Double.parseDouble(data_m[3][15])	; //Transportation fuel consumption (ton/km)
CS3.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][15])	; //Fuel price of transportation (Euro/ton)
CS3.	Electricity_type = data_m[0][17]; //Electricity_type
CS3.	Installation_energy_price	=	Double.parseDouble(data_m[1][17])	; //Energy price (Euro/unit)
CS3. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][15])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
CS3. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][15])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
CS3. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][15])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
CS3. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][15])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
CS3. Spec_GWP_E 	=	Double.parseDouble(data_m[2][17])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
CS3. Spec_AP_E 	=	Double.parseDouble( data_m[3][17]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
CS3. Spec_EP_E 	=	Double.parseDouble( data_m[4][17])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
CS3. Spec_POCP_E 	=	Double.parseDouble(data_m[5][17])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
CS3.H = 7;
CS3.F[0]=Double.parseDouble(data_m1[10][13]);
CS3.F[1]=Double.parseDouble(data_m1[11][13]);
CS3.F[2]=Double.parseDouble(data_m1[12][13]);
CS3.F[3]=Double.parseDouble(data_m1[13][13]);
CS3.F[4]=Double.parseDouble(data_m1[14][13]);
CS3.F[5]=Double.parseDouble(data_m1[15][13]);
CS3.F[6]=Double.parseDouble(data_m1[16][13]);
CS3.C[0]=Double.parseDouble(data_m2[10][13]);
CS3.C[1]=Double.parseDouble(data_m2[11][13]);
CS3.C[2]=Double.parseDouble(data_m2[12][13]);
CS3.C[3]=Double.parseDouble(data_m2[13][13]);
CS3.C[4]=Double.parseDouble(data_m2[14][13]);
CS3.C[5]=Double.parseDouble(data_m2[15][13]);
CS3.C[6]=Double.parseDouble(data_m2[16][13]);
CS3.M[0]=Double.parseDouble(data_m3[10][13]);
CS3.M[1]=Double.parseDouble(data_m3[11][13]);
CS3.M[2]=Double.parseDouble(data_m3[12][13]);
CS3.M[3]=Double.parseDouble(data_m3[13][13]);
CS3.M[4]=Double.parseDouble(data_m3[14][13]);
CS3.M[5]=Double.parseDouble(data_m3[15][13]);
CS3.M[6]=Double.parseDouble(data_m3[16][13]);
CS3.CoTL =CoTL;
CS3.run(); //Run the calculation

//Material: steel, aluminium, composite material...
Parameter_C_Material CM1 = new Parameter_C_Material();
//

CM1.	LOA	=Double.parseDouble(data_m[	1	][5]);
CM1.	LBP	=Double.parseDouble(data_m[	2	][5]);
CM1.	B	=Double.parseDouble(data_m[	3	][5]);	
CM1.	D	=Double.parseDouble(data_m[	4	][5]);
CM1.	T	=Double.parseDouble(data_m[	5	][5]);
CM1.	Cb	=Double.parseDouble(data_m[	6	][5]);
CM1.	Vs	=Double.parseDouble(data_m[	7	][5]);
CM1.	NJ	=Double.parseDouble(data_m[	8	][5]);
CM1.	NE	=Double.parseDouble(data_m[	9	][5]);
CM1.	kA	=Double.parseDouble(data_m[	10	][5]);
CM1.	Pw	=Double.parseDouble(data_m[	11	][5]);

if(Double.parseDouble(data_m[	12	][5])!=0) 
{
	CM1. 	W11 = Double.parseDouble(data_m[	12	][5]);
}

CM1.	E_price		= Double.parseDouble(data_m[	1	][17]);
CM1.	Spec_GWP_E 	= Double.parseDouble(data_m[	2	][17]);
CM1.	Spec_AP_E 	= Double.parseDouble(data_m[	3	][17]);
CM1.	Spec_EP_E 	= Double.parseDouble(data_m[	4	][17]);
CM1.	Spec_POCP_E = Double.parseDouble(data_m[	5	][17]);

CM1.	P_Cutting	=Double.parseDouble(data_m[	5	][6]);
CM1.	V_Cutting	=Double.parseDouble(data_m[	2	][6]);
CM1.	L_Cutting	=Double.parseDouble(data_m[	1	][6]);
CM1.	SMC_Cutting	=Double.parseDouble(data_m[	3	][6]);
CM1.	MP_Cutting 	= Double.parseDouble(data_m[	4	][6]);
//CM1.	H_Cutting_Hull	= Double.parseDouble(data_m[	7	][6]);
if(Double.parseDouble(data_m[	7	][6])!=0) 
{
	CM1.	H_Cutting_Hull	= Double.parseDouble(data_m[	7	][6]);
}
else 
{
	CM1.	r_cutting = Double.parseDouble(data_m[	8	][6]);
}

CM1.P_Bending=Double.parseDouble(data_m[	2	][7]);
CM1.V_Bending=Double.parseDouble(data_m[	1	][7]);

if(Double.parseDouble(data_m[	4	][7])!=0) 
{
	CM1.H_Bending_Hull= Double.parseDouble(data_m[	4	][7]);
}
else 
{
	CM1.	r_bending = Double.parseDouble(data_m[	5	][7]);
}

CM1.P_Welding=Double.parseDouble(data_m[	5	][8]);
CM1.V_Welding=Double.parseDouble(data_m[	2	][8]);
CM1.L_Welding=Double.parseDouble(data_m[	1	][8]);
CM1.SMC_Welding=Double.parseDouble(data_m[	3	][8]);
CM1.MP_Welding = Double.parseDouble(data_m[	4	][8]);

if(Double.parseDouble(data_m[	7	][8])!=0) 
{
	CM1.H_Welding_Hull= Double.parseDouble(data_m[	7	][8]);
}
else 
{
	CM1.	r_welding = Double.parseDouble(data_m[	8	][8]);
}

CM1.P_Blasting=Double.parseDouble(data_m[	2	][9]);
CM1.V_Blasting=Double.parseDouble(data_m[	1	][9]);
if(Double.parseDouble(data_m[	4	][9])!=0) 
{
	CM1.H_Blasting_Hull= Double.parseDouble(data_m[	4	][9]);
}
else 
{
	CM1.	r_blasting = Double.parseDouble(data_m[	5	][9]);
}
CM1.P_Painting=Double.parseDouble(data_m[	5	][10]);
CM1.V_Painting=Double.parseDouble(data_m[	2	][10]);
CM1.SMC_Painting=Double.parseDouble(data_m[	3	][10]);
CM1.MP_Painting = Double.parseDouble(data_m[	4	][10]);
if(Double.parseDouble(data_m[	7	][10])!=0) 
{
	CM1.H_Painting_Hull= Double.parseDouble(data_m[	7	][10]);
}
else 
{
	CM1.	r_painting = Double.parseDouble(data_m[	8	][10]);
}

CM1.Transportation_distance=Double.parseDouble(data_m[	1	][11]);
CM1.Transportation_price = Double.parseDouble(data_m[	2	][11]);
CM1.Transportation_SFOC = Double.parseDouble(data_m[	3	][11]);
CM1.Transportation_fuel_price = Double.parseDouble(data_m[	4	][11]);
CM1.Spec_GWP_Trans=Double.parseDouble(data_m[	5	][11]);
CM1.Spec_AP_Trans=Double.parseDouble(data_m[	6	][11]);
CM1.Spec_EP_Trans=Double.parseDouble(data_m[	7	][11]);
CM1.Spec_POCP_Trans=Double.parseDouble(data_m[	8	][11]);

CM1. Labour_fee_cutting =Double.parseDouble(data_m[	6	][6]);
CM1. Labour_fee_bending =Double.parseDouble(data_m[	3	][7]);
CM1. Labour_fee_welding =Double.parseDouble(data_m[	6	][8]);
CM1. Labour_fee_blasting =Double.parseDouble(data_m[	3	][9]);
CM1. Labour_fee_painting =Double.parseDouble(data_m[	6	][10]);

CM1.run(); //Run the calculation


Parameter_C_Material CM2 = new Parameter_C_Material();
//
CM2.	LOA	=Double.parseDouble(data_m[	1	][5]);
CM2.	LBP	=Double.parseDouble(data_m[	2	][5]);
CM2.	B	=Double.parseDouble(data_m[	3	][5]);	
CM2.	D	=Double.parseDouble(data_m[	4	][5]);
CM2.	T	=Double.parseDouble(data_m[	5	][5]);
CM2.	Cb	=Double.parseDouble(data_m[	6	][5]);
CM2.	Vs	=Double.parseDouble(data_m[	7	][5]);
CM2.	NJ	=Double.parseDouble(data_m[	8	][5]);
CM2.	NE	=Double.parseDouble(data_m[	9	][5]);
CM2.	Pw	=Double.parseDouble(data_m[	11	][5]);

CM2.kB=Double.parseDouble(data_m[	1	][18]);
CM2.WB=Double.parseDouble(data_m[	2	][18]);
CM2.E_price=Double.parseDouble(data_m[	1	][17]);
CM2.Spec_GWP_E = Double.parseDouble(data_m[	2	][17]);
CM2.Spec_AP_E = Double.parseDouble(data_m[	3	][17]);
CM2.Spec_EP_E = Double.parseDouble(data_m[	4	][17]);
CM2.Spec_POCP_E = Double.parseDouble(data_m[	5	][17]);

CM2.P_Cutting=Double.parseDouble(data_m[	5	][19]);
CM2.V_Cutting=Double.parseDouble(data_m[	2	][19]);
CM2.L_Cutting=Double.parseDouble(data_m[	1	][19]);
CM2.SMC_Cutting=Double.parseDouble(data_m[	3	][19]);
CM2.MP_Cutting = Double.parseDouble(data_m[	4	][19]);
if(Double.parseDouble(data_m[	7	][19])!=0) 
{
	CM2.H_Cutting_Hull= Double.parseDouble(data_m[	7	][19]);
}
else 
{
	CM2.	r_cutting = Double.parseDouble(data_m[	8	][19]);
}
CM2.P_Welding=Double.parseDouble(data_m[	5	][20]);
CM2.V_Welding=Double.parseDouble(data_m[	2	][20]);
CM2.L_Welding=Double.parseDouble(data_m[	1	][20]);
CM2.SMC_Welding=Double.parseDouble(data_m[	3	][20]);
CM2.MP_Welding = Double.parseDouble(data_m[	4	][20]);
if(Double.parseDouble(data_m[	7	][20])!=0) 
{
	CM2.H_Welding_Hull= Double.parseDouble(data_m[	7	][20]);
}
else {
	CM2.	r_welding = Double.parseDouble(data_m[	8	][20]);
}

CM2.P_Painting=Double.parseDouble(data_m[	5	][21]);
CM2.V_Painting=Double.parseDouble(data_m[	2	][21]);
CM2.L_Painting=Double.parseDouble(data_m[	1	][21]);
CM2.SMC_Painting=Double.parseDouble(data_m[	3	][21]);
CM2.MP_Painting = Double.parseDouble(data_m[	4	][21]);
if(Double.parseDouble(data_m[	7	][21])!=0) 
{
	CM2.H_Painting_Hull= Double.parseDouble(data_m[	7	][21]);
}
else 
{
	CM2.	r_painting = Double.parseDouble(data_m[	8	][21]);
}

CM2. Labour_fee_cutting =Double.parseDouble(data_m[	6	][19]);
CM2. Labour_fee_welding =Double.parseDouble(data_m[	6	][20]);
CM2. Labour_fee_painting =Double.parseDouble(data_m[	6	][21]);

CM2.run(); //Run the calculation

//Retrofitting
//System: ME, Aux, Scrubber...
Parameter_C_System RS1 = new Parameter_C_System();
RS1.	Engine_type	=	data_m[0][30];					//Engine type
RS1.	Number	=	Double.parseDouble(data_m[1][30])	; //Number of items
RS1.	Weight	=	Double.parseDouble(data_m[2][30])	; //Weight of item (ton)
RS1.	Price	=	Double.parseDouble(data_m[3][30])	; //Item price (Euro/ton)
RS1.	Transportation_type = cb_m[32].getName();							//transportation type
RS1.	Transportation_distance	=	Double.parseDouble(data_m[1][32])	; //Distance of item transportation (km)
RS1.	Transportation_fee	=	Double.parseDouble(data_m[2][32])	; //Transportation price per km (Euro/km)
RS1.	Transportation_SFOC	=	Double.parseDouble(data_m[3][32])	; //Transportation fuel consumption (ton/km)
RS1.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][32])	; //Fuel price of transportation (Euro/ton)
RS1.	Electricity_type = cb_m[34].getName(); //Electricity_type
RS1.	Installation_energy_price	=	Double.parseDouble(data_m[1][34])	; //Energy price (Euro/unit)
RS1. 	Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][32])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
RS1. 	Spec_AP_Trans 	=	Double.parseDouble(data_m[6][32])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
RS1. 	Spec_EP_Trans 	=	Double.parseDouble(data_m[7][32])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
RS1. 	Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][32])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
RS1. 	Spec_GWP_E 	=	Double.parseDouble(data_m[2][34])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
RS1. 	Spec_AP_E 	=	Double.parseDouble( data_m[3][34]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
RS1. 	Spec_EP_E 	=	Double.parseDouble( data_m[4][34])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
RS1. 	Spec_POCP_E 	=	Double.parseDouble(data_m[5][34])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
RS1.H = 7;
RS1.F[0]=Double.parseDouble(data_m1[10][30]);
RS1.F[1]=Double.parseDouble(data_m1[11][30]);
RS1.F[2]=Double.parseDouble(data_m1[12][30]);
RS1.F[3]=Double.parseDouble(data_m1[13][30]);
RS1.F[4]=Double.parseDouble(data_m1[14][30]);
RS1.F[5]=Double.parseDouble(data_m1[15][30]);
RS1.F[6]=Double.parseDouble(data_m1[16][30]);
RS1.C[0]=Double.parseDouble(data_m2[10][30]);
RS1.C[1]=Double.parseDouble(data_m2[11][30]);
RS1.C[2]=Double.parseDouble(data_m2[12][30]);
RS1.C[3]=Double.parseDouble(data_m2[13][30]);
RS1.C[4]=Double.parseDouble(data_m2[14][30]);
RS1.C[5]=Double.parseDouble(data_m2[15][30]);
RS1.C[6]=Double.parseDouble(data_m2[16][30]);
RS1.M[0]=Double.parseDouble(data_m3[10][30]);
RS1.M[1]=Double.parseDouble(data_m3[11][30]);
RS1.M[2]=Double.parseDouble(data_m3[12][30]);
RS1.M[3]=Double.parseDouble(data_m3[13][30]);
RS1.M[4]=Double.parseDouble(data_m3[14][30]);
RS1.M[5]=Double.parseDouble(data_m3[15][30]);
RS1.M[6]=Double.parseDouble(data_m3[16][30]);
RS1.CoTL =CoTL;
RS1.run(); //Run the calculation


Parameter_C_System RS2 = new Parameter_C_System();
RS2.	Engine_type = data_m[0][31];					//battery name
RS2.	Number	=	Double.parseDouble(data_m[1][31])	; //Number of items
RS2.	Weight	=	Double.parseDouble(data_m[2][31])	; //Weight of item (ton)
RS2.	Price	=	Double.parseDouble(data_m[3][31])	; //Item price (Euro/ton)
RS2.	Transportation_type = data_m[0][33];							//transportation type
RS2.	Transportation_distance	=	Double.parseDouble(data_m[1][33])	; //Distance of item transportation (km)
RS2.	Transportation_fee	=	Double.parseDouble(data_m[2][33])	; //Transportation fee per km (Euro/km)
RS2.	Transportation_SFOC	=	Double.parseDouble(data_m[3][33])	; //Transportation fuel consumption (ton/km)
RS2.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][33])	; //Fuel price of transportation (Euro/ton)
RS2.	Electricity_type = data_m[0][34]; //Electricity_type
RS2.	Installation_energy_price	=	Double.parseDouble(data_m[1][34])	; //Energy price (Euro/unit)
RS2. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][33])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
RS2. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][33])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
RS2. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][33])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
RS2. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][33])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
RS2. Spec_GWP_E 	=	Double.parseDouble(data_m[2][34])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
RS2. Spec_AP_E 	=	Double.parseDouble( data_m[3][34]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
RS2. Spec_EP_E 	=	Double.parseDouble( data_m[4][34])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
RS2. Spec_POCP_E 	=	Double.parseDouble(data_m[5][34])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
RS2.H = 7;
RS2.F[0]=Double.parseDouble(data_m1[12][31]);
RS2.F[1]=Double.parseDouble(data_m1[13][31]);
RS2.F[2]=Double.parseDouble(data_m1[14][31]);
RS2.F[3]=Double.parseDouble(data_m1[15][31]);
RS2.F[4]=Double.parseDouble(data_m1[16][31]);
RS2.F[5]=Double.parseDouble(data_m1[17][31]);
RS2.F[6]=Double.parseDouble(data_m1[18][31]);
RS2.C[0]=Double.parseDouble(data_m2[12][31]);
RS2.C[1]=Double.parseDouble(data_m2[13][31]);
RS2.C[2]=Double.parseDouble(data_m2[14][31]);
RS2.C[3]=Double.parseDouble(data_m2[15][31]);
RS2.C[4]=Double.parseDouble(data_m2[16][31]);
RS2.C[5]=Double.parseDouble(data_m2[17][31]);
RS2.C[6]=Double.parseDouble(data_m2[18][31]);
RS2.M[0]=Double.parseDouble(data_m3[12][31]);
RS2.M[1]=Double.parseDouble(data_m3[13][31]);
RS2.M[2]=Double.parseDouble(data_m3[14][31]);
RS2.M[3]=Double.parseDouble(data_m3[15][31]);
RS2.M[4]=Double.parseDouble(data_m3[16][31]);
RS2.M[5]=Double.parseDouble(data_m3[17][31]);
RS2.M[6]=Double.parseDouble(data_m3[18][31]);
RS2.CoTL =CoTL;
RS2.run(); //Run the calculation

Parameter_C_System RS3 = new Parameter_C_System();
RS3.	Engine_type = data_m[0][32];					//battery name
RS3.	Number	=	Double.parseDouble(data_m[1][32])	; //Number of items
RS3.	Weight	=	Double.parseDouble(data_m[2][32])	; //Weight of item (ton)
RS3.	Price	=	Double.parseDouble(data_m[3][32])	; //Item price (Euro/ton)
RS3.	Transportation_type = data_m[0][33];							//transportation type
RS3.	Transportation_distance	=	Double.parseDouble(data_m[1][33])	; //Distance of item transportation (km)
RS3.	Transportation_fee	=	Double.parseDouble(data_m[2][33])	; //Transportation fee per km (Euro/km)
RS3.	Transportation_SFOC	=	Double.parseDouble(data_m[3][33])	; //Transportation fuel consumption (ton/km)
RS3.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][33])	; //Fuel price of transportation (Euro/ton)
RS3.	Electricity_type = data_m[0][34]; //Electricity_type
RS3.	Installation_energy_price	=	Double.parseDouble(data_m[1][34])	; //Energy price (Euro/unit)
RS3. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][33])	; //Specific GWP of transportation (ton CO2e/ ton fuel used for transportation)
RS3. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][33])	; //Specific AP of transportation (ton SO2e/ ton fuel used for transportation)
RS3. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][33])	; //Specific EP of transportation (ton PO4e/ ton fuel used for transportation)
RS3. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][33])	; //Specific POCP of transportation (ton C2H6e/ ton fuel used for transportation)
RS3. Spec_GWP_E 	=	Double.parseDouble(data_m[2][34])	; //Specific GWP of energy consumption (ton CO2e/ MJ energy used)
RS3. Spec_AP_E 	=	Double.parseDouble( data_m[3][34]); //Specific AP of energy consumption (ton SO2e/ MJ energy used)
RS3. Spec_EP_E 	=	Double.parseDouble( data_m[4][34])	; //Specific EP of energy consumption (ton PO4e/ MJ energy used)
RS3. Spec_POCP_E 	=	Double.parseDouble(data_m[5][34])	; //Specific POCP of energy consumption (ton C2H6e/ MJ energy used)
RS3.H = 7;
RS3.F[0]=Double.parseDouble(data_m1[12][32]);
RS3.F[1]=Double.parseDouble(data_m1[13][32]);
RS3.F[2]=Double.parseDouble(data_m1[14][32]);
RS3.F[3]=Double.parseDouble(data_m1[15][32]);
RS3.F[4]=Double.parseDouble(data_m1[16][32]);
RS3.F[5]=Double.parseDouble(data_m1[17][32]);
RS3.F[6]=Double.parseDouble(data_m1[18][32]);
RS3.C[0]=Double.parseDouble(data_m2[12][32]);
RS3.C[1]=Double.parseDouble(data_m2[13][32]);
RS3.C[2]=Double.parseDouble(data_m2[14][32]);
RS3.C[3]=Double.parseDouble(data_m2[15][32]);
RS3.C[4]=Double.parseDouble(data_m2[16][32]);
RS3.C[5]=Double.parseDouble(data_m2[17][32]);
RS3.C[6]=Double.parseDouble(data_m2[18][32]);
RS3.M[0]=Double.parseDouble(data_m3[12][32]);
RS3.M[1]=Double.parseDouble(data_m3[13][32]);
RS3.M[2]=Double.parseDouble(data_m3[14][32]);
RS3.M[3]=Double.parseDouble(data_m3[15][32]);
RS3.M[4]=Double.parseDouble(data_m3[16][32]);
RS3.M[5]=Double.parseDouble(data_m3[17][32]);
RS3.M[6]=Double.parseDouble(data_m3[18][32]);
RS3.CoTL =CoTL;
RS3.run(); //Run the calculation

//Material: steel, aluminium, composite material...
Parameter_C_Material RM1 = new Parameter_C_Material();
RM1.W1 = 0;
RM1. 	W11 = Double.parseDouble(data_m[	13	][5]);

RM1.	E_price		= Double.parseDouble(data_m[	1	][34]);
RM1.	Spec_GWP_E 	= Double.parseDouble(data_m[	2	][34]);
RM1.	Spec_AP_E 	= Double.parseDouble(data_m[	3	][34]);
RM1.	Spec_EP_E 	= Double.parseDouble(data_m[	4	][34]);
RM1.	Spec_POCP_E = Double.parseDouble(data_m[	5	][34]);

RM1.	P_Cutting	=Double.parseDouble(data_m[	5	][24]);
RM1.	V_Cutting	=Double.parseDouble(data_m[	2	][24]);
RM1.	L_Cutting	=Double.parseDouble(data_m[	1	][24]);
RM1.	SMC_Cutting	=Double.parseDouble(data_m[	3	][24]);
RM1.	MP_Cutting 	= Double.parseDouble(data_m[	4	][24]);
if(Double.parseDouble(data_m[	7	][24])!=0)
{
	RM1.	H_Cutting_Hull	= Double.parseDouble(data_m[	7	][24]);
}
else 
{
	RM1.	r_cutting = Double.parseDouble(data_m[	8	][24]);
}

RM1.P_Bending=Double.parseDouble(data_m[	2	][25]);
RM1.V_Bending=Double.parseDouble(data_m[	1	][25]);
if(Double.parseDouble(data_m[	4	][25])!=0) {
	RM1.H_Bending_Hull= Double.parseDouble(data_m[	4	][25]);
}
else {
	RM1.	r_bending = Double.parseDouble(data_m[	5	][25]);
}

RM1.P_Welding=Double.parseDouble(data_m[	5	][26]);
RM1.V_Welding=Double.parseDouble(data_m[	2	][26]);
RM1.L_Welding=Double.parseDouble(data_m[	1	][26]);
RM1.SMC_Welding=Double.parseDouble(data_m[	3	][26]);
RM1.MP_Welding = Double.parseDouble(data_m[	4	][26]);

//if the hours of welding is inputed just use this hours to determine energy consumption (E=P*H). 
//If it leave as 0, use speed and quantity to find energy consumption (E=P*L/V) 
if(Double.parseDouble(data_m[	7	][26])!=0) 
{
	RM1.H_Welding_Hull= Double.parseDouble(data_m[	7	][26]);
}
else 
{
	RM1.	r_welding = Double.parseDouble(data_m[	8	][26]);
}

RM1.P_Blasting=Double.parseDouble(data_m[	2	][27]);
RM1.V_Blasting=Double.parseDouble(data_m[	1	][27]);
if(Double.parseDouble(data_m[	4	][27])!=0) 
{
	RM1.H_Blasting_Hull= Double.parseDouble(data_m[	4	][27]);
}
else 
{
	RM1.	r_blasting = Double.parseDouble(data_m[	5	][27]);
}

RM1.P_Painting=Double.parseDouble(data_m[	5	][28]);
RM1.V_Painting=Double.parseDouble(data_m[	2	][28]);
RM1.L_Painting=Double.parseDouble(data_m[	1	][28]);
RM1.SMC_Painting=Double.parseDouble(data_m[	3	][28]);
RM1.MP_Painting = Double.parseDouble(data_m[	4	][28]);
if(Double.parseDouble(data_m[	7	][28])!=0) 
{
	RM1.H_Painting_Hull= Double.parseDouble(data_m[	7	][28]);
}
else 
{
	RM1.	r_painting = Double.parseDouble(data_m[	8	][28]);
}

RM1.Transportation_distance=Double.parseDouble(data_m[	1	][29]);
RM1.Transportation_price = Double.parseDouble(data_m[	2	][29]);
RM1.Transportation_SFOC = Double.parseDouble(data_m[	3	][29]);
RM1.Transportation_fuel_price = Double.parseDouble(data_m[	4	][29]);
RM1.Spec_GWP_Trans=Double.parseDouble(data_m[	5	][29]);
RM1.Spec_AP_Trans=Double.parseDouble(data_m[	6	][29]);
RM1.Spec_EP_Trans=Double.parseDouble(data_m[	7	][29]);
RM1.Spec_POCP_Trans=Double.parseDouble(data_m[	8	][29]);

RM1. Labour_fee_cutting =Double.parseDouble(data_m[	6	][24]);
RM1. Labour_fee_bending =Double.parseDouble(data_m[	3	][25]);
RM1. Labour_fee_welding =Double.parseDouble(data_m[	6	][26]);
RM1. Labour_fee_blasting =Double.parseDouble(data_m[	3	][27]);
RM1. Labour_fee_painting =Double.parseDouble(data_m[	6	][28]);

RM1.run(); //Run the calculation


Parameter_C_Material RM2 = new Parameter_C_Material();

RM2.kB=Double.parseDouble(data_m[	1	][36]);
RM2.WB=Double.parseDouble(data_m[	2	][36]);
RM2.E_price=Double.parseDouble(data_m[	1	][34]);
RM2.Spec_GWP_E = Double.parseDouble(data_m[	2	][34]);
RM2.Spec_AP_E = Double.parseDouble(data_m[	3	][34]);
RM2.Spec_EP_E = Double.parseDouble(data_m[	4	][34]);
RM2.Spec_POCP_E = Double.parseDouble(data_m[	5	][34]);

RM2.P_Cutting=Double.parseDouble(data_m[	5	][37]);
RM2.V_Cutting=Double.parseDouble(data_m[	2	][37]);
RM2.L_Cutting=Double.parseDouble(data_m[	1	][37]);
RM2.SMC_Cutting=Double.parseDouble(data_m[	3	][37]);
RM2.MP_Cutting = Double.parseDouble(data_m[	4	][37]);
if(Double.parseDouble(data_m[	7	][37])!=0) 
{
	RM2.H_Cutting_Hull= Double.parseDouble(data_m[	7	][37]);

}
else 
{
	RM2.	r_cutting = Double.parseDouble(data_m[	8	][37]);
}
RM2.P_Welding=Double.parseDouble(data_m[	5	][38]);
RM2.V_Welding=Double.parseDouble(data_m[	2	][38]);
RM2.L_Welding=Double.parseDouble(data_m[	1	][38]);
RM2.SMC_Welding=Double.parseDouble(data_m[	3	][38]);
RM2.MP_Welding = Double.parseDouble(data_m[	4	][38]);
if(Double.parseDouble(data_m[	7	][38])!=0) 
{
	RM2.H_Welding_Hull= Double.parseDouble(data_m[	7	][38]);

}
else 
{
	RM2.	r_welding = Double.parseDouble(data_m[	8	][38]);
}

RM2.P_Painting=Double.parseDouble(data_m[	5	][39]);
RM2.V_Painting=Double.parseDouble(data_m[	2	][39]);
RM2.L_Painting=Double.parseDouble(data_m[	1	][39]);
RM2.SMC_Painting=Double.parseDouble(data_m[	3	][39]);
RM2.MP_Painting = Double.parseDouble(data_m[	4	][39]);
if(Double.parseDouble(data_m[	7	][39])!=0) 
{
	RM2.H_Painting_Hull= Double.parseDouble(data_m[	7	][39]);
}
else 
{
	RM2.	r_painting = Double.parseDouble(data_m[	8	][39]);
}

RM2. Labour_fee_cutting =Double.parseDouble(data_m[	6	][37]);
RM2. Labour_fee_welding =Double.parseDouble(data_m[	6	][38]);
RM2. Labour_fee_painting =Double.parseDouble(data_m[	6	][39]);

RM2.run(); //Run the calculation


//Operation M1
Parameter_O O1 = new Parameter_O();
O1.	Life_span	=	Life_span	;
O1.	Interest	=	Interest	;
O1.	PV 			=	PV	;
O1.	Number 		=	CS1.Number;
O1. Fuel_type 	= 	data_m[0][45];
O1. LO_type  	= 	data_m[0][46];

O1.	Ohour		=	Double.parseDouble(data_m[1][42]); //Operation hours (h)
O1.	Ohour_2		=	Double.parseDouble(data_m[8][42]); //Operation hours (h)
O1.	Ohour_3		=	Double.parseDouble(data_m[15][42]); //Operation hours (h)

O1. Engine_No	=   Double.parseDouble(data_m[2][42]); //Number of operated engines (h)
O1. Engine_No_2	=   Double.parseDouble(data_m[9][42]); //Number of operated engines (h)
O1. Engine_No_3	=   Double.parseDouble(data_m[16][42]); //Number of operated engines (h)

O1.	Eload		=	Double.parseDouble(data_m[3][42])	; //Power (kW) (Engine output)
O1.	Eload_2		=	Double.parseDouble(data_m[10][42])	; //Power (kW) (Engine output)
O1.	Eload_3		=	Double.parseDouble(data_m[17][42])	; //Power (kW) (Engine output)

O1.	SFOC		=	Double.parseDouble(data_m[4][42])	; //Specific fuel oil consumption (g/kWh)
O1.	SFOC_2		=	Double.parseDouble(data_m[11][42])	; //Specific fuel oil consumption (g/kWh)
O1.	SFOC_3		=	Double.parseDouble(data_m[18][42])	; //Specific fuel oil consumption (g/kWh)


O1. C_Factor = Double.parseDouble(data_m[6][45]); 	//carbon content
O1. S_Factor = Double.parseDouble(data_m[7][45]);	//sulfer content
O1. N_Factor = Double.parseDouble(data_m[8][45]);        //nitrogen content
O1.	Fuel_price	=	Double.parseDouble(data_m[1][45])	; //Fuel oil price (Euro/ton)



O1.	SLOC		=	Double.parseDouble(data_m[5][42])	; //Specific lubricating oil consumption (g/kWh)
O1.	SLOC_2		=	Double.parseDouble(data_m[12][42])	; //Specific lubricating oil consumption (g/kWh)
O1.	SLOC_3		=	Double.parseDouble(data_m[19][42])	; //Specific lubricating oil consumption (g/kWh)

O1.	LO_price	=	Double.parseDouble(data_m[1][46])	; //Lub oil price (Euro/ton)
O1.	Transportation_fee	=	Double.parseDouble(data_m[2][47])	; //Transportation fee per km (Euro/km)
O1.	Transportation_SFOC	=	Double.parseDouble(data_m[3][47])	; //Transportation specification fuel consumption (ton/km)
O1.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][47])	; //Fuel price of transportation (Euro/ton)
O1.	Transportation_distance	=	Double.parseDouble(data_m[1][47])	; //Distance of material transportation (km)

O1. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][47])	;
O1. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][47])	;
O1. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][47])	;
O1. Spec_POCP_Trans =	Double.parseDouble(data_m[8][47])	;
O1. Spec_GWP_FO 	=	Double.parseDouble(data_m[2][45])		;
O1. Spec_AP_FO 		=	Double.parseDouble(data_m[3][45])		;
O1. Spec_EP_FO 		=	Double.parseDouble(data_m[4][45])		;
O1. Spec_POCP_FO 	=	Double.parseDouble(data_m[5][45])		;
O1. Spec_GWP_LO 	=	Double.parseDouble(data_m[2][46])		;
O1. Spec_AP_LO 		=	Double.parseDouble(data_m[3][46])		;
O1. Spec_EP_LO 		=	Double.parseDouble(data_m[4][46])		;
O1. Spec_POCP_LO 	=	Double.parseDouble(data_m[5][46])		;
O1.run(); //Run the calculation

//Operation M2
Parameter_O O2 = new Parameter_O();
O2.	Life_span	=	Life_span	;
O2.	Interest	=	Interest	;
O2.	PV 			=	PV	;
O2.	Number 		=	CS1.Number;
O2. Fuel_type 	= 	data_m[0][45];
O2. LO_type  	= 	data_m[0][46];

O2.	Ohour		=	Double.parseDouble(data_m[1][43]); //Operation hours (h)
O2.	Ohour_2		=	Double.parseDouble(data_m[8][43]); //Operation hours (h)
O2.	Ohour_3		=	Double.parseDouble(data_m[15][43]); //Operation hours (h)

O2. Engine_No	=   Double.parseDouble(data_m[2][43]); //Number of operated engines (h)
O2. Engine_No_2	=   Double.parseDouble(data_m[9][43]); //Number of operated engines (h)
O2. Engine_No_3	=   Double.parseDouble(data_m[16][43]); //Number of operated engines (h)

O2.	Eload		=	Double.parseDouble(data_m[3][43])	; //Power (kW) (Engine output)
O2.	Eload_2		=	Double.parseDouble(data_m[10][43])	; //Power (kW) (Engine output)
O2.	Eload_3		=	Double.parseDouble(data_m[17][43])	; //Power (kW) (Engine output)

O2.	SFOC		=	Double.parseDouble(data_m[4][43])	; //Specific fuel oil consumption (g/kWh)
O2.	SFOC_2		=	Double.parseDouble(data_m[11][43])	; //Specific fuel oil consumption (g/kWh)
O2.	SFOC_3		=	Double.parseDouble(data_m[18][43])	; //Specific fuel oil consumption (g/kWh)


O2. C_Factor = Double.parseDouble(data_m[6][45]); 	//carbon content
O2. S_Factor = Double.parseDouble(data_m[7][45]);	//sulfer content
O2. N_Factor = Double.parseDouble(data_m[8][45]);        //nitrogen content
O2.	Fuel_price	=	Double.parseDouble(data_m[1][45])	; //Fuel oil price (Euro/ton)



O2.	SLOC		=	Double.parseDouble(data_m[5][43])	; //Specific lubricating oil consumption (g/kWh)
O2.	SLOC_2		=	Double.parseDouble(data_m[12][43])	; //Specific lubricating oil consumption (g/kWh)
O2.	SLOC_3		=	Double.parseDouble(data_m[19][43])	; //Specific lubricating oil consumption (g/kWh)

O2.	LO_price	=	Double.parseDouble(data_m[1][46])	; //Lub oil price (Euro/ton)
O2.	Transportation_fee	=	Double.parseDouble(data_m[2][47])	; //Transportation fee per km (Euro/km)
O2.	Transportation_SFOC	=	Double.parseDouble(data_m[3][47])	; //Transportation specification fuel consumption (ton/km)
O2.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][47])	; //Fuel price of transportation (Euro/ton)
O2.	Transportation_distance	=	Double.parseDouble(data_m[1][47])	; //Distance of material transportation (km)

//Specific GWP of activity (ton CO2e/ ton fuel used for activity)
//Specific AP of activity (ton SO2e/ ton fuel used for activity)
//Specific EP of activity (ton PO4e/ ton fuel used for activity)
//Specific POCP of activity (ton C2H6e/ ton fuel used for activity)

O2. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][47])	;
O2. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][47])	;
O2. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][47])	;
O2. Spec_POCP_Trans =	Double.parseDouble(data_m[8][47])	;
O2. Spec_GWP_FO 	=	Double.parseDouble(data_m[2][45])		;
O2. Spec_AP_FO 		=	Double.parseDouble(data_m[3][45])		;
O2. Spec_EP_FO 		=	Double.parseDouble(data_m[4][45])		;
O2. Spec_POCP_FO 	=	Double.parseDouble(data_m[5][45])		;
O2. Spec_GWP_LO 	=	Double.parseDouble(data_m[2][46])		;
O2. Spec_AP_LO 		=	Double.parseDouble(data_m[3][46])		;
O2. Spec_EP_LO 		=	Double.parseDouble(data_m[4][46])		;
O2. Spec_POCP_LO 	=	Double.parseDouble(data_m[5][46])		;
O2.run(); //Run the calculation

//Operation M3
Parameter_O O3 = new Parameter_O();
O3.	Life_span	=	Life_span	;
O3.	Interest	=	Interest	;
O3.	PV 			=	PV	;
O3.	Number 		=	Double.parseDouble(data_m[2][44]);
O3. Fuel_type 	= 	data_m[0][45];
O3. LO_type  	= 	data_m[0][46];

if(Double.parseDouble(data_m[	22	][44])==1) {
	O3.	SFOC		=	0	; //Specific fuel oil consumption (g/kWh)
	O3.	SFOC_2		=	0	; //Specific fuel oil consumption (g/kWh)
	O3.	SFOC_3		=	0	; //Specific fuel oil consumption (g/kWh)
	}
else {
	O3.	SFOC		=	Double.parseDouble(data_m[4][44])	; //Specific fuel oil consumption (g/kWh)
	O3.	SFOC_2		=	Double.parseDouble(data_m[11][44])	; //Specific fuel oil consumption (g/kWh)
	O3.	SFOC_3		=	Double.parseDouble(data_m[18][44])	; //Specific fuel oil consumption (g/kWh)
}
//System.out.println("sfoc=" + O3.	SFOC);
O3.	Ohour		=	Double.parseDouble(data_m[1][44]); //Operation hours (h)
O3.	Ohour_2		=	Double.parseDouble(data_m[8][44]); //Operation hours (h)
O3.	Ohour_3		=	Double.parseDouble(data_m[15][44]); //Operation hours (h)

O3. Engine_No	=   Double.parseDouble(data_m[2][44]); //Number of operated engines (h)
O3. Engine_No_2	=   Double.parseDouble(data_m[9][44]); //Number of operated engines (h)
O3. Engine_No_3	=   Double.parseDouble(data_m[16][44]); //Number of operated engines (h)

O3.	Eload		=	Double.parseDouble(data_m[3][44])	; //Power (kW) (Engine output)
O3.	Eload_2		=	Double.parseDouble(data_m[10][44])	; //Power (kW) (Engine output)
O3.	Eload_3		=	Double.parseDouble(data_m[17][44])	; //Power (kW) (Engine output)

O3. C_Factor = Double.parseDouble(data_m[6][45]); 	//carbon content
O3. S_Factor = Double.parseDouble(data_m[7][45]);	//sulfer content
O3. N_Factor = Double.parseDouble(data_m[8][45]);        //nitrogen content
O3.	Fuel_price	=	Double.parseDouble(data_m[1][45])	; //Fuel oil price (Euro/ton)



O3.	SLOC		=	Double.parseDouble(data_m[5][44])	; //Specific lubricating oil consumption (g/kWh)
O3.	SLOC_2		=	Double.parseDouble(data_m[12][44])	; //Specific lubricating oil consumption (g/kWh)
O3.	SLOC_3		=	Double.parseDouble(data_m[19][44])	; //Specific lubricating oil consumption (g/kWh)

O3.	LO_price	=	Double.parseDouble(data_m[1][46])	; //Lub oil price (Euro/ton)
O3.	Transportation_fee	=	Double.parseDouble(data_m[2][47])	; //Transportation fee per km (Euro/km)
O3.	Transportation_SFOC	=	Double.parseDouble(data_m[3][47])	; //Transportation specification fuel consumption (ton/km)
O3.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][47])	; //Fuel price of transportation (Euro/ton)
O3.	Transportation_distance	=	Double.parseDouble(data_m[1][47])	; //Distance of material transportation (km)

//Specific GWP of activity (ton CO2e/ ton fuel used for activity)
//Specific AP of activity (ton SO2e/ ton fuel used for activity)
//Specific EP of activity (ton PO4e/ ton fuel used for activity)
//Specific POCP of activity (ton C2H6e/ ton fuel used for activity)

O3. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][47])	;
O3. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][47])	;
O3. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][47])	;
O3. Spec_POCP_Trans =	Double.parseDouble(data_m[8][47])	;
O3. Spec_GWP_FO 	=	Double.parseDouble(data_m[2][45])		;
O3. Spec_AP_FO 		=	Double.parseDouble(data_m[3][45])		;
O3. Spec_EP_FO 		=	Double.parseDouble(data_m[4][45])		;
O3. Spec_POCP_FO 	=	Double.parseDouble(data_m[5][45])		;
O3. Spec_GWP_LO 	=	Double.parseDouble(data_m[2][46])		;
O3. Spec_AP_LO 		=	Double.parseDouble(data_m[3][46])		;
O3. Spec_EP_LO 		=	Double.parseDouble(data_m[4][46])		;
O3. Spec_POCP_LO 	=	Double.parseDouble(data_m[5][46])		;
O3.Electricity_price = Double.parseDouble(data_m[1][17])		;
O3.run(); //Run the calculation
//Maintenance

Parameter_M M1 = new Parameter_M();
M1.GE_number=Double.parseDouble(data_m[2][42]);
M1.B_number=Double.parseDouble(data_m[2][44]);

M1.Life_span = Life_span;
M1.GE_Power= Double.parseDouble(data_m[4][42]);
M1.GEM_Price= Double.parseDouble(data_m[2][48]);
M1.B_Power= Double.parseDouble(data_m[3][44]);
M1.BM_Price= Double.parseDouble(data_m[3][48]);

M1.	E_Cutting	=Double.parseDouble(data_m[5][6]);
M1.	E_Bending	=Double.parseDouble(data_m[3][7]);
M1.	E_Welding	=Double.parseDouble(data_m[5][8]);
M1. E_Blasting = Double.parseDouble(data_m[3][9]);
M1.	E_Painting	=Double.parseDouble(data_m[5][10]);

M1.	H_Cutting_Ma	=Double.parseDouble(data_m[8][48]);
M1.	H_Bending_Ma	=Double.parseDouble(data_m[9][48]);
M1.	H_Welding_Ma	=Double.parseDouble(data_m[10][48]);
M1.	H_Painting_Ma	=Double.parseDouble(data_m[11][48]);

M1.	H_Cutting_O	=Double.parseDouble(data_m[1][50]);
M1.	H_Bending_O	=Double.parseDouble(data_m[2][50]);
M1.	H_Welding_O	=Double.parseDouble(data_m[3][50]);
M1.	H_Painting_O	=Double.parseDouble(data_m[4][50]);

M1.	H_Cutting_Hull	=Double.parseDouble(data_m[1][49]);
M1.	H_Bending_Hull	=Double.parseDouble(data_m[2][49]);
M1.	H_Welding_Hull	=Double.parseDouble(data_m[4][49]);
M1.	H_Painting_Hull	=Double.parseDouble(data_m[5][49]);
M1.	H_Blasting_Hull	=Double.parseDouble(data_m[3][49]);

M1. Spec_GWP_E = Double.parseDouble(data_m[2][17]);
M1. Spec_AP_E = Double.parseDouble(data_m[3][17]);
M1. Spec_EP_E = Double.parseDouble(data_m[4][17]);
M1. Spec_POCP_E = Double.parseDouble(data_m[5][17]);

M1.E_Working_hours= Double.parseDouble(data_m[1][42]);
M1.Boiler_Working_hours= Double.parseDouble(data_m[1][43]);
M1.GE_Working_hours= Double.parseDouble(data_m[1][42]);
M1.B_Working_hours= Double.parseDouble(data_m[1][44]);


M1.Labour_fee_cutting = Double.parseDouble(data_m[6][6]);
M1.Labour_fee_bending = Double.parseDouble(data_m[4][7]);
M1.Labour_fee_welding = Double.parseDouble(data_m[6][8]);
M1.Labour_fee_blasting = Double.parseDouble(data_m[4][9]);
M1.Labour_fee_coating = Double.parseDouble(data_m[6][10]);

M1.engine_price = Double.parseDouble(data_m[3][12])	;
double Spare_Cost = 0;
System.out.println(Double.parseDouble(data_m[1][52]));
for (int i=1; i<data_length;i++) {
	Spare_Cost+=Double.parseDouble(data_m[i][52]);
	}
M1.Spare_cost =Spare_Cost;
System.out.println(M1.Spare_cost);

M1.run();
//Scrapping
Parameter_S S1 = new Parameter_S();
if(Double.parseDouble(data_m[12][5])!=0) {
S1.Hull_Weight = Double.parseDouble(data_m[12][5]);}
else {
	S1.Hull_Weight = CM1.W1;
}
S1.	Machinery_Number	=	CS1.Number	; //Number of item for scrapping
S1.	Machinery_Weight	=	CS1.Weight	; //Weight of item (ton)
S1.S_Price_M	=	Double.parseDouble(data_m[6][54])	; //Price of machinery scrapping (Euro/set)
S1.S_Price_H	= 	Double.parseDouble(data_m[1][54]);//Price of hullscrapping (Euro/set)
S1.	Transportation_distance	=	Double.parseDouble(data_m[1][55])	; //Distance of material transportation (km)
S1.	Transportation_fee	=	Double.parseDouble(data_m[2][55])	; //Transportation fee per km (Euro/km)
S1.	Transportation_SFOC	=	Double.parseDouble(data_m[3][55]); //Transportation specification fuel consumption (ton/km)
S1.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][55])	; //Fuel price of transportation (Euro/ton)
S1.	Electricity_price	=	Double.parseDouble(data_m[1][56])	;  //Fuel oil price (Euro/ton) natural gas
S1.Cutting_power=Double.parseDouble(data_m[5][6]);
S1.Cutting_hours = Double.parseDouble(data_m[1][57])+Double.parseDouble(data_m[1][58])+Double.parseDouble(data_m[1][59]);
S1.Cleaning_power=Double.parseDouble(data_m[5][10]);
S1.Cleaning_hours = Double.parseDouble(data_m[2][57])+Double.parseDouble(data_m[2][58])+Double.parseDouble(data_m[2][59]);

S1.	Life_span	=	Life_span	;
S1.	Interest	=	Interest	;
S1.	PV 	=PV;
//Specific GWP of activity (ton CO2e/ ton fuel used for activity)
//Specific AP of activity (ton SO2e/ ton fuel used for activity)
//Specific EP of activity (ton PO4e/ ton fuel used for activity)
//Specific POCP of activity (ton C2H6e/ ton fuel used for activity)

S1. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][55])	;
S1. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][55])	;
S1. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][55])	;
S1. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][55])	;
S1. Spec_GWP_E 	=	Double.parseDouble(data_m[2][56])	;
S1. Spec_AP_E 	=	Double.parseDouble(data_m[3][56])	;
S1. Spec_EP_E 	=	Double.parseDouble(data_m[4][56])	;
S1. Spec_POCP_E =	Double.parseDouble(data_m[5][56])	;
S1.run(); //Run the calculation


Parameter_S S2 = new Parameter_S();
S2.	Machinery_Number	=	CS1.Number	; //Number of item for scrapping
S2.	Machinery_Weight	=	CS1.Weight	; //Weight of item (ton)
S2.S_Price_M	=	Double.parseDouble(data_m[8][54])	; //Price of scrapping (Euro/ton)
S2.	Transportation_distance	=	Double.parseDouble(data_m[1][55])	; //Distance of material transportation (km)
S2.	Transportation_fee	=	Double.parseDouble(data_m[2][55])	; //Transportation price per km (Euro/km)
S2.	Transportation_SFOC	=	Double.parseDouble(data_m[3][55]); //Transportation specification fuel consumption (ton/km)
S2.	Transportation_fuel_price	=	Double.parseDouble(data_m[4][55])	; //Fuel price of transportation (Euro/ton)

S2.	Life_span	=	0	;
S2.	Interest	=	Interest	;
S2.	PV 	=PV;
//Specific GWP of activity (ton CO2e/ ton fuel used for activity)
//Specific AP of activity (ton SO2e/ ton fuel used for activity)
//Specific EP of activity (ton PO4e/ ton fuel used for activity)
//Specific POCP of activity (ton C2H6e/ ton fuel used for activity)

S2. Spec_GWP_Trans 	=	Double.parseDouble(data_m[5][55])	;
S2. Spec_AP_Trans 	=	Double.parseDouble(data_m[6][55])	;
S2. Spec_EP_Trans 	=	Double.parseDouble(data_m[7][55])	;
S2. Spec_POCP_Trans 	=	Double.parseDouble(data_m[8][55])	;
S2. Spec_GWP_E 	=	Double.parseDouble(data_m[2][56])	;
S2. Spec_AP_E 	=	Double.parseDouble(data_m[3][56])	;
S2. Spec_EP_E 	=	Double.parseDouble(data_m[4][56])	;
S2. Spec_POCP_E =	Double.parseDouble(data_m[5][56])	;
S2.run(); //Run the calculation

//Bulk Results
double Total_cost = CM1.Cost_C_Material+CM2.Cost_C_Material+CS1.Cost_C_System + CS2.Cost_C_System+
					RM1.Cost_C_Material+RM2.Cost_C_Material+RS1.Cost_C_System + RS2.Cost_C_System+
					O1.Cost_O+O2.Cost_O+O3.Cost_O+M1.Cost_M+S1.Cost_S+S2.Cost_S; //Total cost (Euro)  	CM1.Cost_C_Material+
double Total_GWP = 	CM1.GWP+CM2.GWP+CS1.GWP+CS2.GWP+
					RM1.GWP+RM2.GWP+RS1.GWP+RS2.GWP+	
					M1.GWP+O1.GWP+O2.GWP+O3.GWP+S1.GWP+S2.GWP; //Total GWP    							CM1.GWP+
double Total_AP = 	CM1.AP+CM2.AP+CS1.AP+CS2.AP+
					RM1.AP+RM2.AP+RS1.AP+RS2.AP+
					M1.AP+O1.AP+O2.AP+O3.AP+S1.AP+S2.AP; //Total AP									CM1.AP+
double Total_EP = 	CM1.EP+CM2.EP+CS1.EP+CS2.EP+
					RM1.EP+RM2.EP+RS1.EP+RS2.EP+
					M1.EP+O1.EP+O2.EP+O3.EP+S1.EP+S2.EP; //Total EP									CM1.EP+
double Total_POCP = CM1.POCP+CM2.POCP+CS1.POCP+CS2.POCP+
					RM1.POCP+RM2.POCP+RS1.POCP+RS2.POCP+
					M1.POCP+O1.POCP+O2.POCP+O3.POCP+S1.POCP+S2.POCP; //Total POCP							CM1.POCP+
Total_RA = CS1.RA+ CS2.RA+RS1.RA+RS2.RA; //Total RA
Total_CRA = Total_RA/1000*CoTL;

//Results breakdown
/*design*/		double sum0 = 0;
/*C_H*/			double sum1 = CM1.Cost_C_Material; 
/*C_M*/			double sum2 = CS1.Cost_C_System+CS2.Cost_C_System ; //+ CM1.Cost_C_Material
/*C_A*/			double sum3 = CM2.Cost_C_Material;
				double sumC = sum1+sum2+sum3;
				/*R_H*/			double sum1r = RM1.Cost_C_Material; 
				/*R_M*/			double sum2r = RS1.Cost_C_System+RS2.Cost_C_System ; //+ CM1.Cost_C_Material
				/*R_A*/			double sum3r = RM2.Cost_C_Material;
								double sumR = sum1r+sum2r+sum3r;
/*O*/			double sum4 = O1.Cost_O+O2.Cost_O+O3.Cost_O;
/*M*/			double sum5 = M1.Cost_M;
/*S*/			double sum6 = S1.Cost_S+S2.Cost_S;
/*Sum*/			sum = Total_cost;
/*design*/		double GWP0 = 0;				
/*C_H*/			double GWP1 = CM1.GWP;
/*C_M*/			double GWP2 = CS1.GWP + CS2.GWP; //+ CM1.Cost_C_Material
/*C_A*/			double GWP3 = CM2.GWP;
				double GWP_C=GWP1+GWP2+GWP3;
				/*R_H*/			double GWP1r = RM1.GWP;
				/*R_M*/			double GWP2r = RS1.GWP + RS2.GWP; //+ CM1.Cost_C_Material
				/*R_A*/			double GWP3r = RM2.GWP;
								double GWP_R=GWP1r+GWP2r+GWP3r;
/*O*/			double GWP4 = O1.GWP+O2.GWP+O3.GWP;
/*M*/			double GWP5 = M1.GWP;
/*S*/			double GWP6 = S1.GWP+S2.GWP;
/*Sum*/			GWP = Total_GWP;
//P_GWP = 24; //Euro per ton
System.out.println(P_GWP + "test");
/*design*/		double AP0 = 0;				
/*C_H*/			double AP1 = CM1.AP;
/*C_M*/			double AP2 = CS1.AP +CS2.AP; //+ CM1.Cost_C_Material
/*C_A*/			double AP3 = CM2.AP;
				double AP_C = AP1+AP2+AP3;
				/*R_H*/			double AP1r = RM1.AP;
				/*R_M*/			double AP2r = RS1.AP +RS2.AP; //+ CM1.Cost_C_Material
				/*R_A*/			double AP3r = RM2.AP;
								double AP_R = AP1r+AP2r+AP3r;
/*O*/			double AP4 = O1.AP+O2.AP+O3.AP;
/*M*/			double AP5 = M1.AP;
/*S*/			double AP6 = S1.AP +S2.AP;
/*AP*/			AP = Total_AP;
//P_AP = 7788;
/*design*/		double EP0 = 0;				
/*C_H*/			double EP1 = CM1.EP;
/*C_M*/			double EP2 = CS1.EP +CS2.EP; //+ CM1.Cost_C_Material
/*C_A*/			double EP3 = CM2.EP;
				double EP_C = EP1+EP2+EP3;
				/*R_H*/			double EP1r = RM1.EP;
				/*R_M*/			double EP2r = RS1.EP +RS2.EP; //+ CM1.Cost_C_Material
				/*R_A*/			double EP3r = RM2.EP;
								double EP_R = EP1r+EP2r+EP3r;
/*O*/			double EP4 = O1.EP+O2.EP+O3.EP;
/*M*/			double EP5 = M1.EP;
/*S*/			double EP6 = S1.EP +S2.EP;
/*EP*/			EP = Total_EP;
P_EP = 0;
/*design*/		double POCP0 = 0;				
/*C_H*/			double POCP1 = CM1.POCP;
/*C_M*/			double POCP2 = CS1.POCP +CS2.POCP; //+ CM1.Cost_C_Material
/*C_A*/			double POCP3 = CM2.POCP;
				double POCP_C=POCP1+POCP2+POCP3;
				/*R_H*/			double POCP1r = RM1.POCP;
				/*R_M*/			double POCP2r = RS1.POCP +RS2.POCP; //+ CM1.Cost_C_Material
				/*R_A*/			double POCP3r = RM2.POCP;
								double POCP_R=POCP1r+POCP2r+POCP3r;
/*O*/			double POCP4 = O1.POCP+O2.POCP+O3.POCP;
/*M*/			double POCP5 = M1.POCP;
/*S*/			double POCP6 = S1.POCP +S2.POCP;
/*POCP*/		POCP = Total_POCP;
P_POCP = 0;
/*design*/		double RA0 = 0;				
/*C_H*/			double RA1 = 0;
/*C_M*/			double RA2 = 0; //+ CM1.Cost_C_Material
/*C_A*/			double RA3 = 0;
/*O*/			double RA4 = CS1.RA +CS2.RA;
/*O_r*/			double RA4r = RS1.RA+RS2.RA;	
/*M*/			double RA5 = 0;
/*S*/			double RA6 = 0;
/*RA*/			RA = Total_RA;

double Total_LCTC = Total_cost+Total_GWP*P_GWP+Total_AP*P_AP+Total_EP*P_EP+Total_POCP*P_POCP+Total_CRA;
//System.out.println("xxx1="+Total_LCTC);
//System.out.println("xxx2="+sum);
q.add(String.valueOf(sum0));
q.add(String.valueOf(sum1));
q.add(String.valueOf(sum2));
q.add(String.valueOf(sum3));
q.add(String.valueOf(sum4));
q.add(String.valueOf(sum5));
q.add(String.valueOf(sum6));
q.add(String.valueOf(sum));


textReport.setText( 	"ShipLCA Report"+"\n"+
        "Case Name: "+ project_name +"\n"+
        "Date: " +dateFormat.format(date) +"\n"+"\n"+
        
        PV_state+"\n"+
        SL_state+"\n"+"\n"+
        
        "Life cycle cost (Euro):"+"\n"+
        "	Construction (Structure): 	" + formatter.format(sum1+sum3) +"\n"+
        "	Construction (Machinery): 	" + formatter.format(sum2) +"\n"+
        "	Retrofitting (Structure): 	" + formatter.format(sum1r+sum3r) +"\n"+
        "	Retrofitting (Machinery): 	" + formatter.format(sum2r) +"\n"+
        "	Operation: 		" + formatter.format(sum4) + "\n"+
        "	Maintenance: 	" + formatter.format(sum5) + "\n"+
        "	Scrapping: 		" + formatter.format(sum6) + "\n"+
        "	Total cost: 		" + formatter.format(sum) + "\n"+"\n"+
        
        "Global Warming Potential (GWP) (ton CO2e):"+"\n"+
        "	Construction: 	" + formatter.format(GWP_C) +"\n"+
        "	Retrofitting: 		" + formatter.format(GWP_R) +"\n"+
        "	Operation: 		" + formatter.format(GWP4) +"\n"+
        "	Maintenance: 	" +formatter.format(GWP5) +"\n"+
        "	Scrapping: 		" + formatter.format(GWP6) +"\n"+
        "	Total GWP: 		" + formatter.format(Total_GWP) +"\n"+"\n"+
        
        "Acidification Potential (AP) (ton SO2e):"+"\n"+
        "	Construction: 	" + formatter.format(AP_C)+"\n"+
        "	Retrofitting:		" + formatter.format(AP_R)+"\n"+
        "	Operation: 		" + formatter.format(AP4) +"\n"+
        "	Maintenance: 	" +formatter.format(AP5) +"\n"+
        "	Scrapping: 		" + formatter.format(AP6) +"\n"+
        "	Total AP: 		" + formatter.format(Total_AP) +"\n"+"\n"+
        
        "Eutrophication Potential(EP) (ton PO4e):"+"\n"+
        "	Construction: 	" + formatter.format(EP_C) +"\n"+
        "	Retrofitting:		" + formatter.format(EP_R) +"\n"+
        "	Operation:		" + formatter.format(EP4) +"\n"+
        "	Maintenance: 	" +formatter.format(EP5) +"\n"+
        "	Scrapping: 		" + formatter.format(EP6) +"\n"+
        "	Total EP: 		" + formatter.format(Total_EP) +"\n"+"\n"+
        
        
        "Photochemical Ozone Creation Potential (POCP) (ton C2H6e):"+"\n"+
        "	Construction: 	" + formatter.format(POCP_C) +"\n"+
        "	Retrofitting:		" + formatter.format(POCP_R) +"\n"+
        "	Operation: 		" + formatter.format(POCP4 )+"\n"+
        "	Maintenance: 	" +formatter.format(POCP5) +"\n"+
        "	Scrapping: 		" + formatter.format(POCP6) +"\n"+
        "	Total POCP: 		" + formatter.format(Total_POCP) +"\n"+"\n"+
        
        "Risk Priority Number (RPN) of Machinery  (RPN):"+"\n"+
        "	Construction: 	" + formatter.format(RA2) +"\n"+
        "	Retrofitting:		" + formatter.format(RA2) +"\n"+
        "	Operation: 		" + formatter.format(RA4+RA4r) +"\n"+
        "	Maintenance: 	" +formatter.format(RA5) +"\n"+
        "	Scrapping: 		" + formatter.format(RA6) +"\n"+
        "	Total RPN 		" + formatter.format(Total_RA) +"\n"+"\n"+
        
        
        "Total results"+"\n"+
        "	Life cycle cost (Euro): 	"+ formatter.format(sum) +"\n"+
        "	GWP 	(ton CO2e): 	"+ formatter.format(Total_GWP) +"\n"+
        "	GWP 	(Euro): 	"+ formatter.format(Total_GWP*P_GWP) +"\n"+
        "	AP 	(ton SO2e): 	" + formatter.format(Total_AP) +"\n"+
        "	AP 	(Euro): 	" + formatter.format(Total_AP*P_AP) +"\n"+
        "	EP 	(ton PO4e): 	" + formatter.format(Total_EP) +"\n"+
        "	EP 	(Euro): 	" + formatter.format(Total_EP*P_EP) +"\n"+
        "	POCP 	(ton C2H6e): " + formatter.format(Total_POCP) +"\n"+
        "	POCP 	(Euro): 	" + formatter.format(Total_POCP*P_POCP) +"\n"+
        "	RPN:  		" + formatter.format(Total_RA) +"\n"+
        "	Risk Cost 	(Euro): 	" + formatter.format(Total_CRA) +"\n"+
        "	Life cycle total cost (Euro): 	"+ formatter.format(Total_LCTC) +"\n"+"\n"+
        "*Note: now there is no credit regulation for EP and POCP so these results are zero."+"\n"
);

//add XY charts
DefaultCategoryDataset dataset1 =
        new DefaultCategoryDataset( );
dataset1.addValue( sum1+sum2+sum3 , "Construction" , "" );
dataset1.addValue( sumR , "Retrofitting" , "" );
dataset1.addValue( sum4 , "Operation" , "" );
dataset1.addValue( sum5 , "Maintenance" , "");
dataset1.addValue( sum6 , "Scrapping" , "");
dataset1.addValue( sum , "Total" , "");
// Generate the graph
JFreeChart chart1 = ChartFactory.createBarChart(
        "Life Cycle Cost - " + project_name, // Title
        "Life Stages", // x-axis Label
        "Cost (Euro)", // y-axis Label
        dataset1, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP1 = new ChartPanel(chart1);
panel_chart1 = new JPanel();
panel_chart1.add(CP1);

DefaultCategoryDataset dataset2 =
        new DefaultCategoryDataset( );
dataset2.addValue( GWP_C , "Construction" , "" );
dataset2.addValue( GWP_R , "Retrofitting" , "" );

dataset2.addValue( GWP4 , "Operation" , "" );
dataset2.addValue( GWP5 , "Maintenance" , "" );
dataset2.addValue( GWP6 , "Scrapping" , "" );
dataset2.addValue( Total_GWP , "Total" , "" );
// Generate the graph
JFreeChart chart2 = ChartFactory.createBarChart(
        "Global Warming Potential - " + project_name, // Title
        "Life Stages", // x-axis Label
        "GWP (ton CO2e)", // y-axis Label
        dataset2, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP2 = new ChartPanel(chart2);
panel_chart2 = new JPanel();
panel_chart2.add(CP2);

DefaultCategoryDataset dataset3 =
        new DefaultCategoryDataset( );
dataset3.addValue( AP_C , "Construction" , "" );
dataset3.addValue( AP_R , "Retrofitting" , "" );

dataset3.addValue( AP4 , "Operation" , "" );
dataset3.addValue( AP5 , "Maintenance" , "" );
dataset3.addValue( AP6 , "Scrapping" , "" );
dataset3.addValue( Total_AP , "Total" , "" );
// Generate the graph
JFreeChart chart3 = ChartFactory.createBarChart(
        "Acidification Potential - " + project_name, // Title
        "Life Stages", // x-axis Label
        "AP (ton SO2e)", // y-axis Label
        dataset3, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP3 = new ChartPanel(chart3);
panel_chart3 = new JPanel();
panel_chart3.add(CP3);

DefaultCategoryDataset dataset4 =
        new DefaultCategoryDataset( );
dataset4.addValue( EP_C , "Construction" , "" );
dataset4.addValue( EP_R , "Retrofitting" , "" );

dataset4.addValue( EP4 , "Operation" , "" );
dataset4.addValue( EP5 , "Maintenance" , "" );
dataset4.addValue( EP6 , "Scrapping" , "" );
dataset4.addValue( Total_EP , "Total" , "" );
// Generate the graph
JFreeChart chart4 = ChartFactory.createBarChart(
        "Eutrophication Potential - " + project_name, // Title
        "Life Stages", // x-axis Label
        "EP (ton PO4e)", // y-axis Label
        dataset4, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP4 = new ChartPanel(chart4);
panel_chart4 = new JPanel();
panel_chart4.add(CP4);

DefaultCategoryDataset dataset5 =
        new DefaultCategoryDataset( );
dataset5.addValue( POCP_C , "Construction" , "" );
dataset5.addValue( POCP_R , "Retrofitting" , "" );

dataset5.addValue( POCP4 , "Operation" , "" );
dataset5.addValue( POCP5 , "Maintenance" , "" );
dataset5.addValue( POCP6 , "Scrapping" , "" );
dataset5.addValue( Total_POCP , "Total" , "" );
// Generate the graph
JFreeChart chart5 = ChartFactory.createBarChart(
        "Photochemical Ozone Creation Potential - " + project_name, // Title
        "Life Stages", // x-axis Label
        "POCP (ton C2H6e)", // y-axis Label
        dataset5, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP5 = new ChartPanel(chart5);
panel_chart5 = new JPanel();
panel_chart5.add(CP5);

DefaultCategoryDataset dataset6 =
        new DefaultCategoryDataset( );
dataset6.addValue( RA2 , "Construction" , "" );
dataset6.addValue( RA4r , "Retrofitting" , "" );

dataset6.addValue( RA4 , "Operation" , "" );
dataset6.addValue( RA5 , "Maintenance" , "" );
dataset6.addValue( RA6 , "Scrapping" , "" );
dataset6.addValue( Total_RA , "Total" , "" );
// Generate the graph
JFreeChart chart6 = ChartFactory.createBarChart(
        "Risk Assessment- RPN - " + project_name, // Title
        "Life Stages", // x-axis Label
        "Risk Priority Number", // y-axis Label
        dataset6, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP6 = new ChartPanel(chart6);
panel_chart6 = new JPanel();
panel_chart6.add(CP6);

DefaultCategoryDataset dataset7 =
        new DefaultCategoryDataset( );
dataset7.addValue( sum1+sum3+sum2+(GWP1+GWP3+GWP2)*P_GWP+(AP1+AP3+AP2)*P_AP+(EP1+EP3+EP2)*P_EP+(POCP1+POCP3+POCP2)*P_POCP+RA2/1000*CoTL, "Construction" , "" );        //construction total cost
dataset7.addValue( sumR+GWP_R*P_GWP+AP_R*P_AP+EP_R*P_EP+POCP_R*P_POCP+RA4r/1000*CoTL , "Retrofitting" , "" );        	//Retrofitting total cost

dataset7.addValue( sum4+GWP4*P_GWP+AP4*P_AP+EP4*P_EP+POCP4*P_POCP+RA4/1000*CoTL , "Operation" , "" );        	//operation total cost
dataset7.addValue( sum5+GWP5*P_GWP+AP5*P_AP+EP5*P_EP+POCP5*P_POCP+RA5/1000*CoTL, "Maintenance", "");																		//maintenance total cost
dataset7.addValue( sum6+GWP6*P_GWP+AP6*P_AP+EP6*P_EP+POCP6*P_POCP+RA6/1000*CoTL , "Scrapping" , "" ); 			//scrapping total cost
dataset7.addValue( Total_LCTC , "Total" , "" );     				//life cycle total cost
// Generate the graph
JFreeChart chart7 = ChartFactory.createBarChart(
        "Total Life Cycle Cost - " + project_name, // Title
        "Life Stages", // x-axis Label
        "Cost (Euro)", // y-axis Label
        dataset7, // Dataset
        PlotOrientation.HORIZONTAL, // Plot Orientation
        true, // Show Legend
        true, // Use tooltips
        false // Configure chart to generate URLs?
);                              ChartPanel CP7 = new ChartPanel(chart7);
panel_chart7 = new JPanel();

panel_chart7.add(CP7);

Frame = new JFrame("Results plots");

tp_plot1= new JTabbedPane();
tp_plot1.addTab("Cost",ii,panel_chart1,"Cost");
tp_plot1.addTab("GWP",ii,panel_chart2,"GWP");
tp_plot1.addTab("AP",ii,panel_chart3,"AP");
tp_plot1.addTab("EP",ii,panel_chart4,"EP");
tp_plot1.addTab("POCP",ii,panel_chart5,"POCP");
tp_plot1.addTab("RPN",ii,panel_chart6,"RPN");
tp_plot1.addTab("Total cost",ii,panel_chart7,"Total cost");
JPanel panel_plot1 = new JPanel();
panel_plot1.setName("Results plots");
tp_plot1.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
tp_plot1.setTabPlacement(JTabbedPane.TOP);
panel_plot1.add(tp_plot1);

Frame= new JFrame("Results plots");
Frame.add(panel_plot1);
Frame.setExtendedState(JFrame.NORMAL);
Frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
Frame.setResizable(false);
Frame.pack();
Frame.setVisible(true);

F_db_name.dispose();



                            } catch (Exception ex) {
                                java.util.logging.Logger.getLogger(Gui16052019.class.getName()).log(Level.SEVERE, null, ex);
                                JFrame F_warn0 = new JFrame("Error!");
								JPanel P_warn0 = new JPanel();
								P_warn0.setLayout(new BorderLayout());
								JLabel Fd_warn0 = new JLabel("Error: please make sure there are no 'BLANK' cell and run again!");
								F_warn0.setSize(500, 80);
								//Fd_warn0.setText("Database saved in:" + "\n"+ cwd+"/db/" + cb_m[j].getName()+"/"+dateFormat.format(d)+".xls");
								P_warn0.add(Fd_warn0,BorderLayout.CENTER);
								F_warn0.setLocation(400, 170);
								F_warn0.add(P_warn0);
								F_warn0.setVisible(true);
								F_warn0.setResizable(false);
                            }
					}
				});
                        }
 });
   	     
	     uploadButton.addActionListener(new ActionListener() {

			
			@SuppressWarnings({ "static-access", "deprecation" })
			public void actionPerformed(ActionEvent arg0) {

		        LoggingSystem.setUp();

				try {
					LCAdataFactory.setUpDataServiceConnection(urlString, projectString);
				} 
				catch (Exception e) 
					{
					JFrame F_warn0 = new JFrame("Error!");
					JPanel P_warn0 = new JPanel();
					P_warn0.setLayout(new BorderLayout());
					JLabel Fd_warn0 = new JLabel("Can't connect to SHIPLYS server!");
					F_warn0.setSize(200, 80);
					P_warn0.add(Fd_warn0,BorderLayout.CENTER);
					F_warn0.setLocation(w_mid, h_mid);
					F_warn0.add(P_warn0);
					F_warn0.setVisible(true);
					F_warn0.setResizable(false); 						         				
					e.printStackTrace();
				}

GUIexample guiExample = new GUIexample();

//Shows the default property values of the engine CatalogueItem
guiExample.showDefaultEngineValues();

//Shows the default property values of the battery CatalogueItem
guiExample.showDefaultBatteryValues();

//Create new engine product component
ProductComponent engine = guiExample.createEngine(data_m[0][12]);

//Replace the default price value of the new engine
String massValue = new MassMeasure(data_m[2][12], NonSI.TON_UK).toString();
LCAdataFactory.setProductComponentsProperty(engine, "mass", massValue);

String priceValue = new MoneyMeasure(data_m[3][12], Currency.EUR).toString();
LCAdataFactory.setProductComponentsProperty(engine, "price", priceValue);

String powerValue = new PowerMeasure(data_m[4][12], SI.KILO(SI.WATT)).toString();
LCAdataFactory.setProductComponentsProperty(engine, "powerRating", powerValue);

String sfocValue = new SpecificConsumptionMeasure(data_m[4][42]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "sfoc", sfocValue);

String slocValue = new SpecificConsumptionMeasure(data_m[5][42]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "sfoc_lo", slocValue);

String rh1Value = Arrays.asList(data_m1[10][12],data_m2[10][12],data_m3[10][12]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "rh1", rh1Value);

String rh2Value = Arrays.asList(data_m1[11][12],data_m2[11][12],data_m3[11][12]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "rh2", rh2Value);

String rh3Value = Arrays.asList(data_m1[12][12],data_m2[12][12],data_m3[12][12]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "rh3", rh3Value);

String rh4Value = Arrays.asList(data_m1[13][12],data_m2[13][12],data_m3[13][12]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "rh4", rh4Value);

String rh5Value = Arrays.asList(data_m1[14][12],data_m2[14][12],data_m3[14][12]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "rh5", rh5Value);

String rh6Value = Arrays.asList(data_m1[15][12],data_m2[15][12],data_m3[15][12]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "rh6", rh6Value);

String rh7Value = Arrays.asList(data_m1[16][12],data_m2[16][12],data_m3[16][12]).toString();
LCAdataFactory.setProductComponentsProperty(engine, "rh7", rh7Value);



//Create new battery product component
ProductComponent battery = guiExample.createBattery(data_m[0][14]);

String massValue1 = new MassMeasure(data_m[2][14], NonSI.TON_UK).toString();
LCAdataFactory.setProductComponentsProperty(battery, "mass", massValue1);

String priceValue1 = new MoneyMeasure(data_m[3][14], Currency.EUR).toString();
LCAdataFactory.setProductComponentsProperty(battery, "price", priceValue1);

String powerValue1 = new PowerMeasure(data_m[3][44], SI.KILO(SI.WATT)).toString();
LCAdataFactory.setProductComponentsProperty(battery, "powerRating", powerValue1);


String rh1Value1 = Arrays.asList(data_m1[12][14],data_m2[12][14],data_m3[12][14]).toString();
LCAdataFactory.setProductComponentsProperty(battery, "rh1", rh1Value1);

String rh2Value1 = Arrays.asList(data_m1[13][14],data_m2[13][14],data_m3[13][14]).toString();
LCAdataFactory.setProductComponentsProperty(battery, "rh2", rh2Value1);

String rh3Value1 = Arrays.asList(data_m1[14][14],data_m2[14][14],data_m3[14][14]).toString();
LCAdataFactory.setProductComponentsProperty(battery, "rh3", rh3Value1);

String rh4Value1 = Arrays.asList(data_m1[15][14],data_m2[15][14],data_m3[15][14]).toString();
LCAdataFactory.setProductComponentsProperty(battery, "rh4", rh4Value1);

String rh5Value1 = Arrays.asList(data_m1[16][14],data_m2[16][14],data_m3[16][14]).toString();
LCAdataFactory.setProductComponentsProperty(battery, "rh5", rh5Value1);

String rh6Value1 = Arrays.asList(data_m1[17][14],data_m2[17][14],data_m3[17][14]).toString();
LCAdataFactory.setProductComponentsProperty(battery, "rh6", rh6Value1);

String rh7Value1 = Arrays.asList(data_m1[18][14],data_m2[18][14],data_m3[18][14]).toString();
LCAdataFactory.setProductComponentsProperty(battery, "rh7", rh7Value1);





//Create Scenario
AnalysisScenario scenario = LCAdataFactory.createAnalysisScenario("lcaAnalysisScenario" + project_name);

//Create Case
Date date_now = new Date();
AnalysisCase analysisCase = LCAdataFactory.createAnalysisCase(scenario, " - Case: "+dateFormat.format(date_now));

//Create General LCA Parameters
ParameterSet lcaGeneralParameterSet = LCAdataFactory.createParameterSet(analysisCase, "GeneralParameters",
        "This set contains the common lca parameters.", "SHIPLYS general LCA parameter definitions");

LCAdataFactory.setLifeSpan(lcaGeneralParameterSet, Life_span);
LCAdataFactory.setPresentValue(lcaGeneralParameterSet,PV);
LCAdataFactory.setInterestRate(lcaGeneralParameterSet, Interest);
LCAdataFactory.setSensitivityLevel(lcaGeneralParameterSet, (int) SL);
LCAdataFactory.setShipTotalPrice(lcaGeneralParameterSet, CoTL);

  //---log general LCA parameters of the analysis case
log.info("******General");
for (KeyValue kv: lcaGeneralParameterSet.getParameters()) 
{
    log.info("Parameter name =" + kv.getKey() + ", Parameter type =" + kv.getTypeX() + ", Parameter value =" + kv
            .getValue() + "\n");
}

//Create Result - A124 cost

EvaluationResult result_124 = LCAdataFactory.createEvaluationResult(analysisCase, "evaluationResult_124");
double LCTC = sum+GWP*P_GWP+AP*P_AP+EP*P_EP+POCP*P_POCP+RA/1000*CoTL;
LCAdataFactory.setLifeCycleCost(result_124, sum);
LCAdataFactory.setLifeCycleTotalCost(result_124, LCTC);
//LCAdataFactory.setGWP(result, GWP);
LCAdataFactory.setGWPCost(result_124, GWP*P_GWP);
System.out.println("xxx1"+GWP*P_GWP);
System.out.println("xxx2"+GWP);
//LCAdataFactory.setAP(result, AP);
LCAdataFactory.setAPCost(result_124, AP*P_AP);
//LCAdataFactory.setEP(result, EP);
LCAdataFactory.setEPCost(result_124, EP*P_EP);
//LCAdataFactory.setPOCP(result, POCP);
LCAdataFactory.setPOCPCost(result_124, POCP*P_POCP);
//LCAdataFactory.setRPN(result, Total_RA);
LCAdataFactory.setRPNCost(result_124, Total_CRA);

//Save
LCAdataFactory.projOM.makePersistent(scenario);
LCAdataFactory.projOM.makePersistent(analysisCase);
LCAdataFactory.projOM.makePersistent(result_124);

POID<ProcessTemplate> activityPOID=  new POID<>("", ProcessTemplate.class.getName(), POID.buildStructuredName("A12","A124"), "");

//create new activity instance
ActivityInstance actInst =  LCAdataFactory.projOM.createInformationObject(ActivityInstance.class, activityPOID
        .getName()+".instance");

//set the plan attribute of activity instance referencing the activity received from meta service
actInst.setPlan(activityPOID);

//TODO: start date could be set by the work flow controller after selecting a specific activity,
//the appropriate software and starting a job
Date startDate = Date.from(Instant.now().minus(5, ChronoUnit.MINUTES));
actInst.setStartDate(DateHelper.formatIsoDate(startDate));

Date endDate = Date.from(Instant.now());
actInst.setEndDate(DateHelper.formatIsoDate(endDate));

LCAdataFactory.projOM.makePersistent(actInst);
LCAdataFactory.projOM.currentTransaction().commit();


//CreateActivityInstance createActivityInstance_124 = new CreateActivityInstance();
//	Session session_124 = null;
//	try {
//		session_124 = createActivityInstance_124.setUpDataServiceConnection(urlString, projectString);
//	} catch (URISyntaxException e1) {
//		// TODO Auto-generated catch block
//		e1.printStackTrace();
//	}
//System.out.println("this shit is "+session_124.toString());
//createActivityInstance_124.create(session_124.getManager(),"A124");



//Create Result - A127 environment
try {
	LCAdataFactory.setUpDataServiceConnection(urlString, projectString);
} 
catch (Exception e) 
	{
	JFrame F_warn0 = new JFrame("Error!");
	JPanel P_warn0 = new JPanel();
	P_warn0.setLayout(new BorderLayout());
	JLabel Fd_warn0 = new JLabel("Can't connect to SHIPLYS server!");
	F_warn0.setSize(200, 80);
	P_warn0.add(Fd_warn0,BorderLayout.CENTER);
	F_warn0.setLocation(w_mid, h_mid);
	F_warn0.add(P_warn0);
	F_warn0.setVisible(true);
	F_warn0.setResizable(false); 						         				
	e.printStackTrace();
}
EvaluationResult result_127 = LCAdataFactory.createEvaluationResult(analysisCase, "evaluationResult_127");
//double LCTC = sum+GWP*P_GWP+AP*P_AP+EP*P_EP+POCP*P_POCP+RA/1000*CoTL;
//LCAdataFactory.setLifeCycleCost(result, sum);
//LCAdataFactory.setLifeCycleTotalCost(result, LCTC);
LCAdataFactory.setGWP(result_127, GWP);
//LCAdataFactory.setGWPCost(result, GWP*P_GWP);
LCAdataFactory.setAP(result_127, AP);
//LCAdataFactory.setAPCost(result, AP*P_AP);
LCAdataFactory.setEP(result_127, EP);
//LCAdataFactory.setEPCost(result, EP*P_EP);
LCAdataFactory.setPOCP(result_127, POCP);
//LCAdataFactory.setPOCPCost(result, POCP*P_POCP);
//LCAdataFactory.setRPN(result, Total_RA);
//LCAdataFactory.setRPNCost(result, Total_CRA);

//Save

LCAdataFactory.projOM.makePersistent(result_127);
activityPOID=  new POID<>("", ProcessTemplate.class.getName(), POID.buildStructuredName("A12","A127"), "");

//create new activity instance
actInst =  LCAdataFactory.projOM.createInformationObject(ActivityInstance.class, activityPOID
      .getName()+".instance");

//set the plan attribute of activity instance referencing the activity received from meta service
actInst.setPlan(activityPOID);

//TODO: start date could be set by the work flow controller after selecting a specific activity,
//the appropriate software and starting a job
startDate = Date.from(Instant.now().minus(5, ChronoUnit.MINUTES));
actInst.setStartDate(DateHelper.formatIsoDate(startDate));

endDate = Date.from(Instant.now());
actInst.setEndDate(DateHelper.formatIsoDate(endDate));

LCAdataFactory.projOM.makePersistent(actInst);
LCAdataFactory.projOM.currentTransaction().commit();

//CreateActivityInstance createActivityInstance_127 = new CreateActivityInstance();
//Session session_127 = null;
//
//try {
//	session_127 = createActivityInstance_127.setUpDataServiceConnection(urlString, projectString);
//} catch (URISyntaxException e) {
//	// TODO Auto-generated catch block
//	e.printStackTrace();
//}
//createActivityInstance_127.create(session_127.getManager(),"A127");



//JFrame F_ud = new JFrame("Uploaded");
//JPanel P_ud = new JPanel();
//P_ud.setLayout(new BorderLayout());
//JLabel Fd_id = new JLabel("Results are saved to SHIPLYS platform!");
//F_ud.setSize(620, 100);
//P_ud.add(Fd_id,BorderLayout.CENTER);
//F_ud.add(P_ud);
//F_ud.setVisible(true);
//F_ud.setResizable(false);
JFrame F_warn0 = new JFrame("Calculation complete!");
JPanel P_warn0 = new JPanel();
P_warn0.setLayout(new BorderLayout());
JLabel Fd_warn0 = new JLabel("Results are saved and uploaded to SHIPLYS server!");
F_warn0.setSize(400, 80);
P_warn0.add(Fd_warn0,BorderLayout.CENTER);
F_warn0.setLocation(w_mid, h_mid);
F_warn0.add(P_warn0);
F_warn0.setVisible(true);
F_warn0.setResizable(false); 
System.out.println("Results are saved to Project@"+"project_name"+"/~AnalysisCase");
	
}
	    	 
	     });
	     
	     panel6 = createPanel("");	//add panel 
		 tp.addTab("Comparison",ii,panel6,"Comparison");
	     tp.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
	     tp.setTabPlacement(JTabbedPane.TOP);
	     panel6.setLayout(new BoxLayout(panel6,BoxLayout.Y_AXIS));
   
	   

		 
	     JPanel panel61 = new JPanel();
	     panel61.setAlignmentX(Component.CENTER_ALIGNMENT);
	     panel61.setLayout(new FlowLayout());
	     JTextField F61 = new JTextField("Please rename for Alternative 1 results");
	     F61.setFont(font_2);
	     
		 	double w21 = screenSize.getWidth()*1160/1680;
		 	double h21 = screenSize.getHeight()*250/1050;
		 	Dimension d21 = new Dimension();
		 	d21.setSize(w21, h21);
	     
	     F61.setPreferredSize(d21);
	     JButton B61 = new JButton("select file");
	     B61.setFont(font_2);
	     
		 	double w22 = screenSize.getWidth()*360/1680;
		 	double h22 = screenSize.getHeight()*250/1050;
		 	Dimension d22 = new Dimension();
		 	d22.setSize(w22, h22);
	     
	     B61.setPreferredSize(d22);
	     B61.addActionListener(new ActionListener(){
			@Override
	 		public void actionPerformed(ActionEvent e) {
	 			JFileChooser chooser = new JFileChooser(cwd);
	 		    int filename = chooser.showOpenDialog(null);
                                if (filename == JFileChooser.APPROVE_OPTION){
	 			File f = chooser.getSelectedFile();
	 			try {
					Workbook 	wb_r1 = Workbook.getWorkbook(f);
					Sheet 	sheet_r = wb_r1.getSheet(0);
					Cell 	c1=sheet_r.getCell(1, 1);
					Cell 	c2=sheet_r.getCell(2, 1);
					Cell 	c3=sheet_r.getCell(3, 1);
					Cell	c4=sheet_r.getCell(4, 1);
					Cell	c5=sheet_r.getCell(5, 1);
					
					String s1 = c1.getContents(); 
					String s2 = c2.getContents(); 
					String s3 = c3.getContents(); 
					String s4 = c4.getContents(); 
					String s5 = c5.getContents(); 
					
					double r1 = Double.parseDouble(s1);
					double r2 = Double.parseDouble(s2);
					double r3 = Double.parseDouble(s3);
					double r4 = Double.parseDouble(s4);
					double r5 = Double.parseDouble(s5);
					
					a_result[0] =r1;
					a_result[1] =r2;
					a_result[2] =r3;
					a_result[3] =r4;
					a_result[4] =r5;
					
					
				     System.out.println(a_result[0]);	 
				     System.out.println(a_result[1]);	
				     System.out.println(a_result[2]);	
				     System.out.println(a_result[3]);	
				     System.out.println(a_result[4]);	


				     F61.setText(f.getName());
					
					
				} catch (BiffException | IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
	}}});
	
	     
	     JPanel panel62 = new JPanel();
	     panel62.setAlignmentX(Component.CENTER_ALIGNMENT);

	     panel62.setLayout(new FlowLayout());
	     JTextField F62 = new JTextField("Please rename for Alternative 2 results");
	     F62.setFont(font_2);
	     
		 	double w1 = screenSize.getWidth()*1160/1680;
		 	double h1 = screenSize.getHeight()*250/1050;
		 	Dimension d1 = new Dimension();
		 	d1.setSize(w1, h1);
	     
	     F62.setPreferredSize(d1);
	     JButton B62 = new JButton("select file");
	     B62.setFont(font_2);
	     
		 	double w3 = screenSize.getWidth()*360/1680;
		 	double h3 = screenSize.getHeight()*250/1050;
		 	Dimension d3 = new Dimension();
		 	d3.setSize(w3, h3);
		 	
	     B62.setPreferredSize(d3);
	     B62.addActionListener(new ActionListener(){
			@Override
	 		public void actionPerformed(ActionEvent e) {
	 			JFileChooser chooser = new JFileChooser(cwd);
	 		    int filename = chooser.showOpenDialog(null);
	 			if (filename == JFileChooser.APPROVE_OPTION){
	 			File f = chooser.getSelectedFile();
 			
	 			try {
					Workbook 	wb_r1 = Workbook.getWorkbook(f);
					Sheet 	sheet_r = wb_r1.getSheet(0);
					Cell 	c1=sheet_r.getCell(1, 1);
					Cell 	c2=sheet_r.getCell(2, 1);
					Cell 	c3=sheet_r.getCell(3, 1);
					Cell	c4=sheet_r.getCell(4, 1);
					Cell	c5=sheet_r.getCell(5, 1);
					
					String s1 = c1.getContents(); 
					String s2 = c2.getContents(); 
					String s3 = c3.getContents(); 
					String s4 = c4.getContents(); 
					String s5 = c5.getContents(); 
					
					double r1 = Double.parseDouble(s1);
					double r2 = Double.parseDouble(s2);
					double r3 = Double.parseDouble(s3);
					double r4 = Double.parseDouble(s4);
					double r5 = Double.parseDouble(s5);
					
					b_result[0] =r1;
					b_result[1] =r2;
					b_result[2] =r3;
					b_result[3] =r4;
					b_result[4] =r5;
					
					
				     System.out.println(b_result[0]);	 
				     System.out.println(b_result[1]);	
				     System.out.println(b_result[2]);	
				     System.out.println(b_result[3]);	
				     System.out.println(b_result[4]);	

				     F62.setText(f.getName());

					
					
				} catch (BiffException | IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
	 			
	 			
	 			
	}}});
	     
	     JPanel panel63 = new JPanel();
	     panel63.setAlignmentX(Component.CENTER_ALIGNMENT);

	     panel63.setLayout(new FlowLayout());
	     JTextField F63 = new JTextField("Please rename for Alternative 3 results");
	     
		 	double w31 = screenSize.getWidth()*1160/1680;
		 	double h31 = screenSize.getHeight()*250/1050;
		 	Dimension d31 = new Dimension();
		 	d31.setSize(w31, h31);
	     
	     F63.setPreferredSize(d31);
	     JButton B63 = new JButton("select file");
	     F63.setFont(font_2);
	     B63.setFont(font_2);

		 	double w32 = screenSize.getWidth()*360/1680;
		 	double h32 = screenSize.getHeight()*250/1050;
		 	Dimension d32 = new Dimension();
		 	d32.setSize(w32, h32);
	     
	     B63.setPreferredSize(d32);
	     B63.addActionListener(new ActionListener(){
			@Override
	 		public void actionPerformed(ActionEvent e) {
	 			JFileChooser chooser = new JFileChooser(cwd);
	 		    int filename = chooser.showOpenDialog(null);
	 			if (filename == JFileChooser.APPROVE_OPTION){
	 			File f = chooser.getSelectedFile();
	 			try {
					Workbook 	wb_r1 = Workbook.getWorkbook(f);
					Sheet 	sheet_r = wb_r1.getSheet(0);
					Cell 	c1=sheet_r.getCell(1, 1);
					Cell 	c2=sheet_r.getCell(2, 1);
					Cell 	c3=sheet_r.getCell(3, 1);
					Cell	c4=sheet_r.getCell(4, 1);
					Cell	c5=sheet_r.getCell(5, 1);
					
					String s1 = c1.getContents(); 
					String s2 = c2.getContents(); 
					String s3 = c3.getContents(); 
					String s4 = c4.getContents(); 
					String s5 = c5.getContents(); 
					
					double r1 = Double.parseDouble(s1);
					double r2 = Double.parseDouble(s2);
					double r3 = Double.parseDouble(s3);
					double r4 = Double.parseDouble(s4);
					double r5 = Double.parseDouble(s5);
					
					c_result[0] =r1;
					c_result[1] =r2;
					c_result[2] =r3;
					c_result[3] =r4;
					c_result[4] =r5;
					
					
				     System.out.println(c_result[0]);	 
				     System.out.println(c_result[1]);	
				     System.out.println(c_result[2]);	
				     System.out.println(c_result[3]);	
				     System.out.println(c_result[4]);	

				     F63.setText(f.getName());

					
					
					
				} catch (BiffException | IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
	 			
	 			
	 			
	}}}); 
	     
	     panel61.add(F61,LEFT_ALIGNMENT);
	     panel61.add(B61,LEFT_ALIGNMENT);
	     panel62.add(F62,LEFT_ALIGNMENT);
	     panel62.add(B62,LEFT_ALIGNMENT);
	     panel63.add(F63,LEFT_ALIGNMENT);
	     panel63.add(B63,LEFT_ALIGNMENT);

	    panel6.add(panel61);
	    panel6.add(panel62);
	    panel6.add(panel63);


	    		 
	     
	     JButton Button_FC = new JButton("Compare Results");
	     Button_FC.setName("Compare Results");
	     Button_FC.setAlignmentX(Component.CENTER_ALIGNMENT);

	     panel6.add(Button_FC);
	     Button_FC.addActionListener(new ActionListener(){
				
		@Override
		public void actionPerformed(ActionEvent e) {
	    		 
				    DefaultCategoryDataset dataset_r = 	
				    	      new DefaultCategoryDataset( );
				      dataset_r.addValue( a_result[0], F61.getText(),"Construction"); 
				      dataset_r.addValue( b_result[0], F62.getText(),"Construction");
				      dataset_r.addValue( c_result[0], F63.getText(),"Construction");        	

				      dataset_r.addValue( a_result[1], F61.getText(),"Operation");      
				      dataset_r.addValue( b_result[1], F62.getText(),"Operation");
				      dataset_r.addValue( c_result[1], F63.getText(),"Operation"); 
				      
				      dataset_r.addValue( a_result[2], F61.getText(),"Maintenance");       	
				      dataset_r.addValue( b_result[2], F62.getText(),"Maintenance");
				      dataset_r.addValue( c_result[2], F63.getText(),"Maintenance"); 
 
				      dataset_r.addValue( a_result[3], F61.getText(),"Scrapping"); 	
				      dataset_r.addValue( b_result[3], F62.getText(),"Scrapping"); 
				      dataset_r.addValue( c_result[3], F63.getText(),"Scrapping"); 

				      dataset_r.addValue( a_result[4], F61.getText(),"Total");       	
				      dataset_r.addValue( b_result[4], F62.getText(),"Total"); 
				      dataset_r.addValue( c_result[4], F63.getText(),"Total"); 

				      
					
				  // Generate the graph	
				     JFreeChart chart_r = ChartFactory.createBarChart(	
				     "Comparison of Total Life Cycle Cost", // Title	
				     "Life Stages", // x-axis Label	
				     "Total Life Cycle Cost (Euro)", // y-axis Label	
				     dataset_r, // Dataset	
				     PlotOrientation.HORIZONTAL, // Plot Orientation	
				     true, // Show Legend	
				     true, // Use tooltips	
				     false // Configure chart to generate URLs?	
				  );	
				     ChartPanel CP_r = new ChartPanel(chart_r);	
				     	
				     panel_chart_r = new JPanel();	
				     panel_chart_r.add(CP_r);	
				     //create a window for piechart	
				     frame_r = new JFrame("Comparison of Total Life Cycle Cost");	
				     frame_r.setLayout(new BorderLayout());	
				     frame_r.add(panel_chart_r, BorderLayout.WEST);	
				     frame_r.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);	
				     frame_r.setSize(new Dimension(800, 600));	
				     frame_r.setResizable(true);	
				     
					 	double w1 = screenSize.getWidth()*720/1680;
					 	double h1 = screenSize.getHeight()*440/1050;
					 	Dimension d1 = new Dimension();
					 	d1.setSize(w1, h1);
				     
				     panel_chart_r.setPreferredSize(d1);	
				     frame_r.pack();	
				     frame_r.setVisible(true);	
				    
				    
			}});
	     
	 
	}
//create main panel	
	private JPanel createPanel(String string) {
	     
	     panel = new JPanel(true);
	     
		 	double w = screenSize.getWidth()*1480/1680;
		 	double h = screenSize.getHeight()*840/1050;
		 	Dimension d = new Dimension();
		 	d.setSize(w, h);
	     
	     panel.setPreferredSize(d);
	     panel.setLayout(new BorderLayout(1,1));
	     JLabel filler = new JLabel(string);
	     filler.setFont(font_2);
	     filler.setHorizontalAlignment(JLabel.CENTER);
	     panel.add(filler);
	     return panel;
	  }
	
	private JPanel createWelcomePanel(String string) {
	     
		panel_n = createPanel(string);
	 	panel_n.setVisible(true);
	 	
	 	double w0 = screenSize.getWidth()*1200/1680;
	 	double h0 = screenSize.getHeight()*125/1050;
	 	Dimension d0 = new Dimension();
	 	d0.setSize(w0, h0);
	 	
	 	panel_n.setPreferredSize(d0);
	    panel_n.setLayout(new GridBagLayout()); // added code
	    GridBagConstraints c_sub = new GridBagConstraints();
	    //Dimension d = new Dimension(100,80);

//create label
	    JLabel lbl = new JLabel();
	    lbl.setName(string);
	    lbl.setFont(font_3);
	    c_sub.fill = GridBagConstraints.VERTICAL;
	    c_sub.weighty = 1;
	    c_sub.gridx = 0;
	    c_sub.gridy = 0;
	    panel_n.add(lbl,c_sub);
	    
	    area = new JTextArea(string);
	    area.setFont(font_1);
	    c_sub.fill = GridBagConstraints.VERTICAL;
	    c_sub.weighty = 1;
	    c_sub.gridx = 0;
	    c_sub.gridy = 1;
	    area.setEditable(false);
	    area.setAutoscrolls(true);

	    panel_n.add(area,c_sub);
	    panel_n.setVisible(true); // added code
	    return panel_n; 
	  }
//create subpanels with lables, drop lists and text field 	
	private JPanel createsubPanel(String string) throws BiffException, IOException {

			panel_m = createPanel(string);
		 	panel_m.setVisible(true);
		 	double w = screenSize.getWidth()*520/1680;
		 	double h = screenSize.getHeight()*380/1050;
		 	Dimension d = new Dimension();
		 	d.setSize(w, h);
		 	panel_m.setPreferredSize(d);
		    panel_m.setLayout(new GridBagLayout()); // added code
		    GridBagConstraints c_sub = new GridBagConstraints();
		    
		    //create label
		    JLabel lbl = new JLabel();
		    lbl.setName(string);
		    lbl.setFont(font_2);
		    
		 	double w1 = screenSize.getWidth()*520/1680;
		 	double h1 = screenSize.getHeight()*30/1050;
		 	Dimension d1 = new Dimension();
		 	d1.setSize(w1, h1);
		    
		    lbl.setPreferredSize(d1);
		    
		    c_sub.fill = GridBagConstraints.VERTICAL;
		    c_sub.weighty = 1;
		    c_sub.gridx = 0;
		    c_sub.gridy = 0;
		    panel_m.add(lbl,c_sub);

//create drop list for database selection

		    file = new File (cwd+"/db/"+string+"/");//.substring(0, position)
		    System.out.println(cwd);
		    file_group = file.listFiles();
		    Arrays.sort(file_group, Comparator.comparingLong(File::lastModified));
		    choices = new String[file_group.length];
		    choices_NAME = new String[file_group.length];
		    int[] pos = new int[file_group.length];
		    for(int j=0;j<file_group.length;j++){
		    	  
		    	choices_NAME[j]= file_group[j].getName();
		    	pos[j]=choices_NAME[j].lastIndexOf(".");// get rid of extention name e.g. "xls"
		    	if(pos[j] >0){
		    		choices[j]=choices_NAME[j].substring(0, pos[j]);
		    	}}

		    cb = new JComboBox<String>(choices);
			cb.setName(string);
		    cb.setBackground(Color.white);
		    cb.setFont(font_1);
//		    cb.setEditable(true);
		    cb.setSelectedIndex(-1);
		    cb.setSelectedItem("-select database-");
		 	double w11 = screenSize.getWidth()*400/1680;
		 	double h11 = screenSize.getHeight()*35/1050;
		 	Dimension d11 = new Dimension();
		 	d11.setSize(w11, h11);
		    cb.setPreferredSize(d11);
		    		    
		    ((JLabel)cb.getRenderer()).setHorizontalAlignment(SwingConstants.CENTER);
		    
//		    if(cb.getName().contains("Transportation")) {
//		    	((JLabel)cb.getRenderer()).setHorizontalAlignment(SwingConstants.LEFT);
//		    }
		    
//		    cb_update= new ActionListener() {
//				
//				@Override
//				public void actionPerformed(ActionEvent e) {
//					cb.removeAll();
//				    panel_m.add(cb,c_sub);
//				}
//			};
//		    	
//		    
//		    cb.addActionListener(cb_update);
// added code
		    c_sub.fill = GridBagConstraints.VERTICAL;
		    c_sub.weighty = 1;
		    c_sub.gridx = 0;
		    c_sub.gridy = 1;
		    panel_m.add(cb,c_sub);

//create drop list for database selection
		    importData = new JButton("Import from platform");
		    importData.setName("import from server");
		    importData.setToolTipText("SHIPLYS platform should be running before import!");
		    importData.setFont(font_1);
		    
		 	double w2 = screenSize.getWidth()*400/1680;
		 	double h2 = screenSize.getHeight()*20/1050;
		 	Dimension d2 = new Dimension();
		 	d2.setSize(w2, h2);
		 	
		    importData.setPreferredSize(d2);
		    importData.setBackground(Color.white);
		    c_sub.fill = GridBagConstraints.VERTICAL;
		    c_sub.weighty = 1;
		    c_sub.gridx = 0;
		    c_sub.gridy = 3;
//		    importData.addActionListener(createImportListener());
		    panel_m.add(importData,c_sub);
		    
		    
		    
//create field for user inputs		    		    
		    field = new JTextField("0");
		    
		 	double w3 = screenSize.getWidth()*400/1680;
		 	double h3 = screenSize.getHeight()*40/1050;
		 	Dimension d3 = new Dimension();
		 	d3.setSize(w3, h3);
		 	
		    field.setPreferredSize(d3);
		    field.setHorizontalAlignment(JTextField.CENTER);
		    field.setBackground(Color.lightGray);
		    field.setSelectedTextColor(Color.white);
		    field.setSelectionColor(Color.red);
		    field.setFont(font_1);
		    
		    c_sub.fill = GridBagConstraints.VERTICAL;
		    c_sub.weighty = 1;
		    c_sub.gridx = 0;
		    c_sub.gridy = 2;
		    panel_m.add(field,c_sub);

		    panel_m.setVisible(true); // added code
		    return panel_m; 
			
 }	
	
////create a search panel	
//	private JPanel createSearchPanel(String string) throws BiffException, IOException {
//		
//		panel_s = new JPanel();
//	 	panel_s.setVisible(true);
//	 	panel_s.setPreferredSize(new Dimension(200, 100));
//	    panel_s.setLayout(new GridBagLayout()); // added code
//	    GridBagConstraints c_s = new GridBagConstraints();
//
//	    //create label
//	    JLabel lbl = new JLabel("Online database");
//	    lbl.setName("Online database");
//	    lbl.setFont(font_1);
//	    lbl.setPreferredSize(new Dimension(400,60));
//	    c_s.fill = GridBagConstraints.VERTICAL;
//	    c_s.weighty = 0;
//	    c_s.gridx = 0;
//	    c_s.gridy = 0;
//	    panel_s.add(lbl,c_s);	    
//	  //create field for user inputs		    		    
//	    field_search = new JTextField("");
//	    field_search.setName("");
//	    field_search.setPreferredSize(new Dimension(300,60));
//	    field_search.setFont(font_1);
//	    c_s.fill = GridBagConstraints.HORIZONTAL;
//	    c_s.weighty = 0;
//	    c_s.gridx = 0;
//	    c_s.gridy = 1;
//	    panel_s.add(field_search,c_s);
//	    
//	  //create a button for database search
//	    button_search = new JButton("Search");
//	    button_search.setName("Search");
//	    button_search.setFont(font_1);
//	    
//	 	double w = screenSize.getWidth()*100/1680;
//	 	double h = screenSize.getHeight()*60/1050;
//	 	Dimension d = new Dimension();
//	 	d.setSize(w, h);
//	    
//	    button_search.setPreferredSize(d);
//	    c_s.fill = GridBagConstraints.HORIZONTAL;
//	    c_s.weighty = 0;
//	    c_s.gridx = 1;
//	    c_s.gridy = 1;
//	    button_search.addActionListener(createSearchListener(0));
//	    panel_s.add(button_search,c_s);
//
//	    return panel_s;	
//	}
//add icons	
	 @SuppressWarnings("unused")
	private ImageIcon createImageIcon(String string) {
		 string = cwd+"/pic/icon.png";
	     if(string == null)
	     {
	         System.out.println("the image "+string+" is not exist!");
	         return null;
	     }
	     else return new ImageIcon(string);
	  }
	 
	 
//add action ddbb
	 private ActionListener createCbActionListener(int j)
	 {
		 ActionListener ddbb=  new ActionListener() 
		 {
                                @Override
				public void actionPerformed(ActionEvent db) 
                                {
                       	
				selection_Number = Integer.toString(cb_m[j].getSelectedIndex());
				selection_Name = cb_m[j].getItemAt(cb_m[j].getSelectedIndex());
				cb_m[j].setToolTipText(selection_Name);
				System.out.println(cb_m[j].getName());
				cb_m[j].setBackground(Color.red);
				cb_m[j].setForeground(Color.white);
				if(cb_m[j].getName().startsWith("1.")||cb_m[j].getName().startsWith("3.") ||cb_m[j].getName().startsWith("4.")) {
					cb_m[j].setBackground(Color.white);
					cb_m[j].setForeground(Color.black);		
					}
//				||cb_m[j].getName().contains(arg0))
				System.out.println(file_group.length);
				file_1 = new File (cwd+"/db/"+cb_m[j].getName()+"/");
				System.out.println(cwd+"/db/"+cb_m[j].getName());
	    		file_group_1 =  file_1.listFiles();
	    		   		 
	    		Arrays.sort(file_group_1, Comparator.comparingLong(File::lastModified));
	    		  wb = new Workbook[file_group_1.length];
	    		  
	    			sheet = new Sheet[file_group_1.length];
	    			item0  = new Cell[file_group_1.length][data_length];
	    			cell0  = new Cell[file_group_1.length][data_length];
	    			cell1  = new Cell[file_group_1.length][data_length];
	    			cell2  = new Cell[file_group_1.length][data_length];
	    			unit0  = new Cell[file_group_1.length][data_length];
	    			content0  = new String[file_group_1.length][data_length];
	    			content1  = new String[file_group_1.length][data_length];
	    			content2  = new String[file_group_1.length][data_length];
	    			content3  = new String[file_group_1.length][data_length];
	    			content4  = new String[file_group_1.length][data_length];
	    		 
	    		 System.out.println(file_group_1.length);
					data = new Object[data_length][5];
					Object[] columnNames = 
						{
							"Type	", "Average", "Minimum","Maximum","Unit",
		                   };
				
	    		 for (i=0; i<file_group_1.length;i++){

          
						try {
							wb[i] = Workbook.getWorkbook(file_group_1[i]);
						} catch (BiffException | IOException e1) {
							JFrame F_warn0 = new JFrame("Error!");
							JPanel P_warn0 = new JPanel();
							P_warn0.setLayout(new BorderLayout());
							JLabel Fd_warn0 = new JLabel("No database found!");
							F_warn0.setSize(200, 80);
							P_warn0.add(Fd_warn0,BorderLayout.CENTER);
							F_warn0.setLocation(w_mid, h_mid);
							F_warn0.add(P_warn0);
							F_warn0.setVisible(true);
							F_warn0.setResizable(false); 								e1.printStackTrace();
						}
						sheet[i]  = wb[i].getSheet(0);
	    				for(int j = 0 ; j<data_length ; j++)
	    				{
	    					item0[i][j]=sheet[i].getCell(1, j);
	    					cell0[i][j]=sheet[i].getCell(2, j);
	    					cell1[i][j]=sheet[i].getCell(3, j);
	    					cell2[i][j]=sheet[i].getCell(4, j);
	    					unit0[i][j]=sheet[i].getCell(5, j);
	    					
	    					content0[i][j]=item0	[i][j].getContents();
	    					content1[i][j]=cell0	[i][j].getContents();
	    					content2[i][j]=cell1	[i][j].getContents();
	    					content3[i][j]=cell2	[i][j].getContents();
	    					content4[i][j]=unit0	[i][j].getContents();
	    				}

						
							for(int k=0;k<data_length;k++){
								if(Double.parseDouble(selection_Number)==i)
								{
//									System.out.println(i);
								data0[k]	 = 	content0	[i][k];
								data1[k]	 = 	content1	[i][k];
								data2[k]	 = 	content2	[i][k];
								data3[k]	 = 	content3	[i][k];
								data4[k]	 = 	content4	[i][k];}
								
//								if(Double.parseDouble(selection_Number)==1){
//									System.out.println("xxxx");
//									}
								data[k][0] = data0[k];
								data[k][1] = data1[k];
								data[k][2] = data2[k];
								data[k][3] = data3[k];
								data[k][4] = data4[k];
								
								
								
								DTM = new DefaultTableModel(data, columnNames) {
									 public boolean isCellEditable(int row, int column) {
										 if(column == 1||column == 2||column == 3){

											 return true;

											 }else{

											 return false;

											 }						                }
								};
																
								}}
	    		 
	    		 TableWithOverwrite[] table_db1 = new TableWithOverwrite[activity_length];
				 table_db1 [j] = new TableWithOverwrite(DTM);
				 table_db1 [j].setSelectionBackground(Color.GREEN);
				 table_db1 [j].setToolTipText("Please press 'ENTER' after changing/adding!");
				 table_db1 [j].setColumnSelectionAllowed(true);
				 
				 		 
				 JTextField textField = new JTextField();
				 textField.setFont(font_1);
				 textField.setBackground(Color.red);
				 textField.setForeground(Color.white);
				 DefaultCellEditor dce = new DefaultCellEditor( textField );

				 table_db1[j].getColumnModel().getColumn(1).setCellEditor(dce);
				 table_db1[j].getColumnModel().getColumn(2).setCellEditor(dce);
				 table_db1[j].getColumnModel().getColumn(3).setCellEditor(dce);
				 
				 table_db1 [j].getColumnModel().getColumn(0).setPreferredWidth(480);
				 table_db1 [j].getColumnModel().getColumn(1).setPreferredWidth(120);
				 table_db1 [j].getColumnModel().getColumn(2).setPreferredWidth(120);
				 table_db1 [j].getColumnModel().getColumn(3).setPreferredWidth(120);
				 table_db1 [j].getColumnModel().getColumn(4).setPreferredWidth(120);
//				 Font font_head = new Font("Arial",Font.BOLD,14);
				 table_db1 [j].getTableHeader().setFont(font_1);
				 
				 DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
				 centerRenderer.setHorizontalAlignment( JLabel.CENTER );
				 for(int i =0; i<5;i++) {
				 table_db1 [j].getColumnModel().getColumn(i).setCellRenderer(centerRenderer);
				 }
				 
			     JScrollPane sp_db1 = new JScrollPane(table_db1[j]);
			    
				 	double w = screenSize.getWidth()*1000/1680;
				 	double h = screenSize.getHeight()*420/1050;
				 	Dimension d = new Dimension();
				 	d.setSize(w, h);
			     
			     sp_db1.setPreferredSize(d);
			     
	     	     JLabel label_db1 = new JLabel("Database selection");
			     
	     	 	double w1 = screenSize.getWidth()*20/1680;
			 	double h1 = screenSize.getHeight()*20/1050;
			 	Dimension d1 = new Dimension();
			 	d1.setSize(w1, h1);
	     	     
	     	     label_db1.setSize(d1);
    
				 panel_db = new JPanel();
			     panel_db.setName("database"+selection_Number);
			     label_db1.setAlignmentX(Component.CENTER_ALIGNMENT);
			     panel_db.add(sp_db1,BorderLayout.PAGE_START);
			     table_db1[j].setAlignmentX(Component.CENTER_ALIGNMENT);
		
			   //add button	  
			     JButton[] updateButton = new JButton[activity_length];
			     updateButton[j] = new JButton("Update/Save");
			     panel_db.add(updateButton[j],BorderLayout.LINE_START);

		//add action for clicking button	     
			     
			     ActionListener update = new ActionListener() {
					

					
					@Override
					public void actionPerformed(ActionEvent e) {
						
						TableModel DTM2 = table_db1[j].getModel();
						
						JFrame F_db_name = new JFrame("Name database");
						JPanel P_db_name = new JPanel();
						
						F_db_name.setLocation(w_mid, h_mid);
						F_db_name.setSize(new Dimension(350,120));
						JTextField TF_db_name = new JTextField("new database name");
						TF_db_name.setFont(font_1);
						TF_db_name.setPreferredSize(new Dimension(200,50));
						JButton save_name = new JButton("Update & save");
						JButton cancel_name = new JButton("Update only");
						save_name.setFont(font_1);
						cancel_name.setFont(font_1);
						save_name.setPreferredSize(new Dimension(175,20));
						cancel_name.setPreferredSize(new Dimension(150,20));
						
						P_db_name.setLayout(new BorderLayout());
						P_db_name.add(TF_db_name,BorderLayout.PAGE_START);
						P_db_name.add(save_name,BorderLayout.WEST);
						P_db_name.add(cancel_name,BorderLayout.EAST);
						
						F_db_name.add(P_db_name);
						F_db_name.setVisible(true);
						F_db_name.setResizable(false); 
						
						
						ActionListener cancel_name1 = new ActionListener() {
							
							@Override
							public void actionPerformed(ActionEvent cancel_name) {
								
								for (int k = 0;k<data_length;k++){
									data_m0[k][j]=(String) DTM2.getValueAt(k, 0);
									data_m1[k][j]=(String) DTM2.getValueAt(k, 1);
									data_m2[k][j]=(String) DTM2.getValueAt(k, 2);
									data_m3[k][j]=(String) DTM2.getValueAt(k, 3);
									int expression = Integer.parseInt(field0[3].getText());
									switch(expression) {
									  case 1:
										  if(j!=0&&k!=0&&data_m0[k][j]!=""&&(data_m1[k][j]!=""||data_m3[k][j]!="")) {
											  if(data_m2[k][j]=="") {
												JFrame F_warn0 = new JFrame("Error!");
												JPanel P_warn0 = new JPanel();
												P_warn0.setLayout(new BorderLayout());
												JLabel Fd_warn0 = new JLabel(data_m0[k][j]+" data is empty! Please fill in and update again!");
												F_warn0.setSize(500, 80);
												P_warn0.add(Fd_warn0,BorderLayout.CENTER);
												F_warn0.setLocation(400, 170);
												F_warn0.add(P_warn0);
												F_warn0.setVisible(true);
												F_warn0.setResizable(false);
															}}
										  break;
									  case 2:
										  if(j!=0&&k!=0&&data_m0[k][j]!=""&&(data_m2[k][j]!=""||data_m1[k][j]!="")) {
											  if(data_m3[k][j]=="") {
												JFrame F_warn0 = new JFrame("Error!");
												JPanel P_warn0 = new JPanel();
												P_warn0.setLayout(new BorderLayout());
												JLabel Fd_warn0 = new JLabel(data_m0[k][j]+" data is empty! Please fill in and update again!");
												F_warn0.setSize(500, 80);
												P_warn0.add(Fd_warn0,BorderLayout.CENTER);
												F_warn0.setLocation(400, 170);
												F_warn0.add(P_warn0);
												F_warn0.setVisible(true);
												F_warn0.setResizable(false);}}
										  break;
									  default:
										  if(j!=0&&k!=0&&data_m0[k][j]!=""&&(data_m2[k][j]!=""||data_m3[k][j]!="")) {
											  if(data_m1[k][j]=="") {
												JFrame F_warn0 = new JFrame("Error!");
												JPanel P_warn0 = new JPanel();
												P_warn0.setLayout(new BorderLayout());
												JLabel Fd_warn0 = new JLabel(data_m0[k][j]+" data is empty! Please fill in and update again!");
												F_warn0.setSize(500, 80);
												P_warn0.add(Fd_warn0,BorderLayout.CENTER);
												F_warn0.setLocation(400, 170);
												F_warn0.add(P_warn0);
												F_warn0.setVisible(true);
												F_warn0.setResizable(false);
												
					}}}}
								//
								
								F_db_name.dispose();
								frame_db.dispose();
								cb_m[j].setBackground(Color.DARK_GRAY);
								cb_m[j].setForeground(Color.white);	

							}
						};
						cancel_name.addActionListener(cancel_name1);
						
						ActionListener save_name1 = new ActionListener() {
							
							@Override
							public void actionPerformed(ActionEvent save_name) {


								for (int k = 0;k<data_length;k++){
									data_m0[k][j]=(String) DTM2.getValueAt(k, 0);
									data_m1[k][j]=(String) DTM2.getValueAt(k, 1);
									data_m2[k][j]=(String) DTM2.getValueAt(k, 2);
									data_m3[k][j]=(String) DTM2.getValueAt(k, 3);
									int expression = Integer.parseInt(field0[3].getText());
									switch(expression) {
									  case 1:
										  if(j!=0&&k!=0&&data_m0[k][j]!=""&&(data_m1[k][j]!=""||data_m3[k][j]!="")) {
											  if(data_m2[k][j]=="") {
												JFrame F_warn0 = new JFrame("Error!");
												JPanel P_warn0 = new JPanel();
												P_warn0.setLayout(new BorderLayout());
												JLabel Fd_warn0 = new JLabel(data_m0[k][j]+" data is empty! Please fill in and update again!");
												F_warn0.setSize(500, 80);
												P_warn0.add(Fd_warn0,BorderLayout.CENTER);
												F_warn0.setLocation(400, 170);
												F_warn0.add(P_warn0);
												F_warn0.setVisible(true);
												F_warn0.setResizable(false);
															}}
										  break;
									  case 2:
										  if(j!=0&&k!=0&&data_m0[k][j]!=""&&(data_m2[k][j]!=""||data_m1[k][j]!="")) {
											  if(data_m3[k][j]=="") {
												JFrame F_warn0 = new JFrame("Error!");
												JPanel P_warn0 = new JPanel();
												P_warn0.setLayout(new BorderLayout());
												JLabel Fd_warn0 = new JLabel(data_m0[k][j]+" data is empty! Please fill in and update again!");
												F_warn0.setSize(500, 80);
												P_warn0.add(Fd_warn0,BorderLayout.CENTER);
												F_warn0.setLocation(400, 170);
												F_warn0.add(P_warn0);
												F_warn0.setVisible(true);
												F_warn0.setResizable(false);}}
										  break;
									  default:
										  if(j!=0&&k!=0&&data_m0[k][j]!=""&&(data_m2[k][j]!=""||data_m3[k][j]!="")) {
											  if(data_m1[k][j]=="") {
												JFrame F_warn0 = new JFrame("Error!");
												JPanel P_warn0 = new JPanel();
												P_warn0.setLayout(new BorderLayout());
												JLabel Fd_warn0 = new JLabel(data_m0[k][j]+" data is empty! Please fill in and update again!");
												F_warn0.setSize(500, 80);
												P_warn0.add(Fd_warn0,BorderLayout.CENTER);
												F_warn0.setLocation(400, 170);
												F_warn0.add(P_warn0);
												F_warn0.setVisible(true);
												F_warn0.setResizable(false);
												
					}}
											
									}
								
									
							
								}
								//
							
						FileOutputStream fout;
						@SuppressWarnings("resource")
						HSSFWorkbook wb_save = new HSSFWorkbook();
						HSSFRow[] c0 = new HSSFRow[data_length];
						HSSFCell [] Column00 = new HSSFCell[data_length];
						HSSFCell [] Column0 = new HSSFCell[data_length];
						HSSFCell [] Column1 = new HSSFCell[data_length];
						HSSFCell [] Column2 = new HSSFCell[data_length];
						HSSFCell [] Column3 = new HSSFCell[data_length];
						HSSFCell [] Column4 = new HSSFCell[data_length];
						HSSFSheet sheet1 = wb_save.createSheet();
						
						for (int k = 0;k<data_length;k++){
							data_m0[k][j]=(String) DTM2.getValueAt(k, 0);
							data_m1[k][j]=(String) DTM2.getValueAt(k, 1);
							data_m2[k][j]=(String) DTM2.getValueAt(k, 2);
							data_m3[k][j]=(String) DTM2.getValueAt(k, 3);
							data_m4[k][j]=(String) DTM2.getValueAt(k, 4);
							
							c0[k]=sheet1.createRow(k);
							Column00 [k]=c0[k].createCell(0);
							Column0 [k]=c0[k].createCell(1);
							Column1 [k]=c0[k].createCell(2);
							Column2 [k]=c0[k].createCell(3);
							Column3 [k]=c0[k].createCell(4);
							Column4 [k]=c0[k].createCell(5);
							Column00 [k].setCellValue(k);
							Column0 [k].setCellValue(data_m0[k][j]);
							Column1 [k].setCellValue(data_m1[k][j]);
							Column2 [k].setCellValue(data_m2[k][j]);
							Column3 [k].setCellValue(data_m3[k][j]);
							Column4 [k].setCellValue(data_m4[k][j]);
					
							try {
								fout = new FileOutputStream(cwd+"/db/" + cb_m[j].getName()+"/"+ TF_db_name.getText()+".xls");
								wb_save.write(fout);
							} catch (IOException e1) {
								JFrame F_warn0 = new JFrame("Error!");
								JPanel P_warn0 = new JPanel();
								P_warn0.setLayout(new BorderLayout());
								JLabel Fd_warn0 = new JLabel("Can't save to local drive!");
								F_warn0.setSize(200, 80);
								P_warn0.add(Fd_warn0,BorderLayout.CENTER);
								F_warn0.setLocation(w_mid, h_mid);
								F_warn0.add(P_warn0);
								F_warn0.setVisible(true);
								F_warn0.setResizable(false); 									e1.printStackTrace();
							}
						}
						
						JFrame F_warn0 = new JFrame("Upated and saved");
						JPanel P_warn0 = new JPanel();
						P_warn0.setLayout(new BorderLayout());
						JLabel Fd_warn0 = new JLabel("Database saved in:" + "\n"+ cwd+"/db/" + cb_m[j].getName()+"/"+ TF_db_name.getText()+".xls");
						F_warn0.setSize(620, 100);
						F_warn0.setLocation(w_mid, h_mid);
						P_warn0.add(Fd_warn0,BorderLayout.CENTER);
						F_warn0.add(P_warn0);
						F_warn0.setVisible(true);
						F_warn0.setResizable(false);
//add item to array	
					
						cb_m[j].addItem(TF_db_name.getText());
						
//						cb_m[j].setBackground(Color.DARK_GRAY);
//						cb_m[j].setForeground(Color.white);
			 
							frame_db.dispose();
							F_db_name.dispose();
							cb_m[j].setSelectedIndex(cb_m[j].getItemCount()-1);
							cb_m[j].setBackground(Color.DARK_GRAY);
							cb_m[j].setForeground(Color.white);
							frame_db.dispose();
							
								
							
							}
						};
						save_name.addActionListener(save_name1);
						
					}}  ;
					
					
					
//					saveButton[j].addActionListener(saveAsAction);
//					updateButton[j].addActionListener(saveAsAction);
			     //create a window for db
					updateButton[j].addActionListener(update);
			     frame_db = new JFrame(cb_m[j].getName() +": "+selection_Name);//cb_m[j].getName()
			     frame_db.setLayout(new BorderLayout());
			     frame_db.add(panel_db,BorderLayout.WEST);
			     frame_db.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
			     
				 	double w2 = screenSize.getWidth()*1000/1680;
				 	double h2 = screenSize.getHeight()*865/1050;
				 	Dimension d2 = new Dimension();
				 	d2.setSize(w2, h2);
			     
			     frame_db.setSize(d2);
			     frame_db.setResizable(false);
			     
				 	double w21 = screenSize.getWidth()*1000/1680;
				 	double h21 = screenSize.getHeight()*465/1050;
				 	Dimension d21 = new Dimension();
				 	d21.setSize(w21, h21);
			     
			     panel_db.setPreferredSize(d21);
			     frame_db.pack();
			   
			     if(cb_m[j].getSelectedItem().toString().contains("-select database-")) {
			    	 frame_db.setVisible(false);
			     }
			     //hide "-select database-"
//			     if(Double.parseDouble(selection_Number)==0){frame_db.setVisible(false);} 
			     else{
			    	 frame_db.setVisible(true);
			    	}
			      }					
				};
	return ddbb;
	 }
	 


	 public class TableWithOverwrite extends JTable {
		 /**
		 * 
		 */
		private static final long serialVersionUID = 1L;

		public final static String EXCLUDE = "F2";

		 private boolean isBlankEditor = false;

		 public TableWithOverwrite() {
		 super();
		 }

		 public TableWithOverwrite(TableModel tm) {
		 super(tm);
		 }

		 @Override
		 public Component prepareEditor(TableCellEditor editor, int row, int column) {
		 Component c = super.prepareEditor(editor, row, column);

		 if (isBlankEditor)
		 ((JTextField) c).setText("");

		 return c;
		 }

		 @Override
		 protected boolean processKeyBinding(KeyStroke ks, KeyEvent e, int condition, boolean pressed) {
		 if (! EXCLUDE.equals(KeyEvent.getKeyText(e.getKeyCode())))
		 isBlankEditor = true;

		 boolean retValue = super.processKeyBinding(ks, e, condition, pressed);

		 isBlankEditor = false;
		 return retValue;
		 }
		 }

	//add actionlistner to importdata button to open a window for downloading and maping data from server to local database
	 private ActionListener createImportListener(){
		 ActionListener importLinster=  new ActionListener() {
             @Override
             public void actionPerformed(ActionEvent iL) {
            	           	 
            	JFrame F_id = new JFrame("Import from Platform");
             	F_id.setIconImage(new ImageIcon(cwd+"/pic/icon.png").getImage());
         		JPanel P_id = new JPanel();
         		P_id.setLayout(new BorderLayout());
         		JLabel TF = new JLabel("<html><br/>Download available database from SHIPLYS platform?<br/>Make sure SHIPLYS server is running!<br/></html>", SwingConstants.CENTER);
         		
        	 	double w = screenSize.getWidth()*400/1680;
			 	double h = screenSize.getHeight()*120/1050;
			 	Dimension d = new Dimension();
			 	d.setSize(w, h);
         		
         		TF.setPreferredSize(d);
//         		TF.setText("Do you want to download availabel database from SHIPLYS platform? "
//         				+"\n"+ "Make sure SHIPLYS server is running!");
         		JButton download = new JButton();
         		download.setText("Download");
         		download.setFont(font_1);
         		
         		
        	 	double w1 = screenSize.getWidth()*140/1680;
			 	double h1 = screenSize.getHeight()*5/1050;
			 	Dimension d1 = new Dimension();
			 	d1.setSize(w1, h1);
         		
         		download.setPreferredSize(d1);
//         		TF.setBackground(Color.green);
//         		download.setBackground(Color.red);
//         		download.setForeground(Color.white);
//         		
         		JButton cancel = new JButton();

         		cancel.setText("Cancel");
         		cancel.setFont(font_1);
         		
        	 	double w2 = screenSize.getWidth()*140/1680;
			 	double h2 = screenSize.getHeight()*5/1050;
			 	Dimension d2 = new Dimension();
			 	d2.setSize(w2, h2);
         		
         		cancel.setPreferredSize(d2);
         		ActionListener c = new ActionListener() {
         			
         			@Override
         			public void actionPerformed(ActionEvent e) {
         				F_id.dispose();

         			}
         			};
				cancel.addActionListener(c);
//         		cancel.setBackground(Color.red);
//         		cancel.setForeground(Color.white);

         		ActionListener en = new ActionListener() {
         			
         			@Override
         			public void actionPerformed(ActionEvent e) {
         				frame_name = TF.getText();
         				LoggingSystem.setUp();
						 try {
							LCAdataFactory.setUpDataServiceConnection(urlString, projectString);
						} catch (Exception e1) {
							JFrame F_warn0 = new JFrame("Error!");
							JPanel P_warn0 = new JPanel();
							P_warn0.setLayout(new BorderLayout());
							JLabel Fd_warn0 = new JLabel("No SHIPLYS server found!");
							F_warn0.setSize(200, 80);
							P_warn0.add(Fd_warn0,BorderLayout.CENTER);
							F_warn0.setLocation(w_mid, h_mid);
							F_warn0.add(P_warn0);
							F_warn0.setVisible(true);
							F_warn0.setResizable(false); 						         				
							F_id.dispose();
						}

							readFromDB();
							shutdown();
         				F_id.dispose();
         			}
         		};
            	 
         		download.addActionListener(en);
//        		TF.addActionListener(en);
         		
        	 	double w3 = screenSize.getWidth()*400/1680;
			 	double h3 = screenSize.getHeight()*180/1050;
			 	Dimension d3 = new Dimension();
			 	d3.setSize(w3, h3);
         		
        		F_id.setSize(d3);
        		P_id.add(TF,BorderLayout.NORTH);
        		P_id.add(download,BorderLayout.WEST);
        		P_id.add(cancel,BorderLayout.EAST);
        		F_id.add(P_id);
        		F_id.setVisible(true);
        		F_id.setResizable(false);  

        		F_id.setLocation(w_mid,h_mid);
        		F_id.setAlwaysOnTop(true);

             }
	 };	
	 return importLinster;
	 }

	protected void readFromDB() {
		
		//--TODO: read from database, convert data and save it localy
        //E.g.: read all engines and batties from data base like it is shown in example code (see following example and explanations in QueryExamplesGeneral)
        Query query = LCAdataFactory.projOM.newQuery(ProductComponent.class);
        query.setMode(Query.Mode.PERSISTENT);
        @SuppressWarnings("unchecked")
		List<POID<ProductComponent>> result = (List<POID<ProductComponent>>) query.execute();
        log.info("Found " + result.size() + " Product Components:");
        for (POID<ProductComponent> poid: result) {
            ProductComponent productComponent = LCAdataFactory.projOM.getObjectById(poid);

            if (productComponent.getCatalogueItemRef().getCatalogue().getCommonName().equals(LCAdataFactory.catalogueName)) {

                log.info("*******" + productComponent.getCommonName());
            //1. Read values of configurable properties contained in parameterSet of a specific product component (e.g. price etc.)
            ParameterSet parameterSet = productComponent.getParameters();

            List<KeyValue> configurableProperties = null;
          String[] CP_name = new String[54];
          String[] import_name = new String[54];

          String[] CP_value = new String[54];
          String[] import_value = new String[54];
            log.info("Configurable Properties of " + productComponent.getCommonName() + " are:");
            if (parameterSet != null) {
                configurableProperties = parameterSet.getParameters(); 
//                System.out.println(configurableProperties.size());
                CP_name= new String[configurableProperties.size()];
                CP_value= new String[configurableProperties.size()];
                import_name= new String[configurableProperties.size()];
                import_value= new String[configurableProperties.size()];
                for (KeyValue engingeProductComponentProp: configurableProperties) {
                    log.info("Parameter name =" + engingeProductComponentProp.getKey() + ", Parameter type ="
                            + engingeProductComponentProp.getTypeX() + ", Parameter value =" + engingeProductComponentProp
                            .getValue() + "\n");
             }
                             
              for(int i=0;i<configurableProperties.size();i++){
           	   CP_name[i]=configurableProperties.get(i).getKey(); //parameter names
           	   CP_value[i]=configurableProperties.get(i).getValue(); //parameter values
           	   import_name[i]=CP_name[i].substring(CP_name[i].lastIndexOf(".")+1,CP_name[i].length());
            }

            JTable table_id = new JTable();
            table_id  = new JTable(DTM);
            
				TableModel DTM2 = table_id.getModel();
				FileOutputStream fout1;
				@SuppressWarnings("resource")
				HSSFWorkbook wb_save = new HSSFWorkbook();
				HSSFRow[] c0 = new HSSFRow[data_length];
				HSSFCell [] Column00 = new HSSFCell[data_length];
				HSSFCell [] Column0 = new HSSFCell[data_length];
				HSSFCell [] Column1 = new HSSFCell[data_length];

				HSSFSheet sheet1 = wb_save.createSheet();
//				System.out.println(configurableProperties.size());
				for (int k = 0;k<configurableProperties.size();k++){
					
					c0[k]=sheet1.createRow(k);
					Column00 [k]=c0[k].createCell(0);
					Column0 [k]=c0[k].createCell(1);
					Column1 [k]=c0[k].createCell(2);

					Column00 [k].setCellValue(k);
					Column0 [k].setCellValue(import_name[k]);
					Column1 [k].setCellValue(CP_value[k]);
		
					
					try {
						fout1 = new FileOutputStream(cwd+"/db/" + "import database"+"/"+productComponent.getCommonName()+"_configuratable_"+dateFormat.format(date)+".xls");
						wb_save.write(fout1);
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}}}
            
//2. Then read standard values (which are not configurable) from refered catalogueItem
            Set<KeyValue> allProperties = productComponent.getCatalogueItemRef().getProperties();
            KeyValue[] a = allProperties.toArray(new KeyValue[0]);
            for(int i = 0; i<a.length;i++) {
//            System.out.println("test " +i+". "+ a[i].getKey());}
            String[] test= new String[54];
            log.info("Standard Properties of " + productComponent.getCommonName() + " are:");
            for (KeyValue catalogueItemProp: allProperties) {
                CatalogueProperty property = getCataloguePropertyByName(productComponent.getCatalogueItemRef(),
                        catalogueItemProp.getKey());
                if (property.getTypeX() != CataloguePropertyType.CONFIGURABLE) {
                    log.info("Parameter name =" + catalogueItemProp.getKey() + ", Parameter type =" + catalogueItemProp
                            .getTypeX() + ", Parameter value =" + catalogueItemProp.getValue() + "\n");
              }
            }
            JTable table_id = new JTable();
            table_id  = new JTable(DTM);
            
				TableModel DTM3 = table_id.getModel();
				FileOutputStream fout2;
				@SuppressWarnings("resource")
				HSSFWorkbook wb_save2 = new HSSFWorkbook();
				HSSFRow[] c0 = new HSSFRow[data_length];
				HSSFCell [] Column00 = new HSSFCell[data_length];
				HSSFCell [] Column0 = new HSSFCell[data_length];
				HSSFCell [] Column1 = new HSSFCell[data_length];

				HSSFSheet sheet1 = wb_save2.createSheet();
//				System.out.println(a.length);
				for (int k = 0;k<a.length;k++){
					
					c0[k]=sheet1.createRow(k);
					Column00 [k]=c0[k].createCell(0);
					Column0 [k]=c0[k].createCell(1);
					Column1 [k]=c0[k].createCell(2);

					Column00 [k].setCellValue(a[k].getKey());
					Column0 [k].setCellValue(a[k].getValue());
					
					try {
						fout2 = new FileOutputStream(cwd+"/db/" + "import database"+"/"+productComponent.getCommonName()+"_unconfiguratable_"+dateFormat.format(date)+".xls");
						wb_save2.write(fout2);
					} catch (IOException e1) {
						JFrame F_warn0 = new JFrame("Error!");
						JPanel P_warn0 = new JPanel();
						P_warn0.setLayout(new BorderLayout());
						JLabel Fd_warn0 = new JLabel("Can't download to local drive!");
						F_warn0.setSize(200, 80);
						P_warn0.add(Fd_warn0,BorderLayout.CENTER);
						F_warn0.setLocation(w_mid, h_mid);
						F_warn0.add(P_warn0);
						F_warn0.setVisible(true);
						F_warn0.setResizable(false); 							e1.printStackTrace();
					}

					
				}
        
					}
			JFrame F_id = new JFrame("Saved");
			JPanel P_id = new JPanel();
			P_id.setLayout(new BorderLayout());
			JLabel Fd_id = new JLabel("<html>Database saved in:" + "<br/>"+ 	cwd+"/db/" + "import database"+"/"+productComponent.getCommonName()+"_configuratable_"+dateFormat.format(date)+".xls"+"<br/>"+
																	cwd+"/db/" + "import database"+"/"+productComponent.getCommonName()+"_unconfiguratable_"+dateFormat.format(date)+".xls<html>", SwingConstants.CENTER);
			
			F_id.setSize(720, 100);
			P_id.add(Fd_id,BorderLayout.CENTER);
			F_id.add(P_id);
			F_id.setVisible(true);
			F_id.setResizable(false);  
            	}		
        	}
        }
    	
    private static void shutdown()
    {
        if (session != null) {
            session.close();
        }

        InformationDirectory.shutdown(null);
    }
	
	private static CatalogueProperty getCataloguePropertyByName(CatalogueItemInstance item, String propertyName)
    {
        //Load catalogue
        CatalogueManager catM = CatalogueManager.getInstance(LCAdataFactory.projOM);

        Set<POID<Catalogue>> catPOIDs = catM.list();
        if (CollectionsHelper.isNullOrEmpty(catPOIDs)) {
            throw new ObjectNotFoundException("no catalogues defined in context");
        }
        
        catM.open(catPOIDs.stream().filter(catPoid -> catPoid.getName().equals(LCAdataFactory.catalogueName))
                .findFirst().orElse(null));

        CatalogueProperty property = item.getAllPropertyDefinitions().stream()
                .filter(prop -> prop.getCommonName().equals(propertyName))
                .findFirst().get();

        return property;
    }


		 private ActionListener createSearchListener(int j){
			 ActionListener searchListener=  new ActionListener() {


				@Override
	             public void actionPerformed(ActionEvent iL) {
	            	Object[] columnNames = {
							"Number", "Name","Details"
		                   };
	            	JFrame frame_odb = new JFrame("Database found");
	            	String[] result_name= new String[result_size];
	            	String[] test_name = new String[]{"aaa","bbb","ccc","ddd","eee"};
	            	Object[][] data_list = new Object[result_size][3];
	            	checkBox = new JCheckBox[result_size];
	            	JPanel checkPanel = new JPanel();
	            	checkPanel.scrollRectToVisible(getVisibleRect());
	            	JScrollPane checkScrollPanel = new JScrollPane(checkPanel);
	            	checkScrollPanel.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
	            	checkScrollPanel.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
	                
	                GridLayout cpLayout = new GridLayout(0,(int) Math.floor((result_size-1)/25+1));
	                cpLayout.setHgap(1);
	                cpLayout.setVgap(1);
	                cpLayout.minimumLayoutSize(frame_odb);
	            	checkPanel.setLayout(cpLayout);
	            	
	        	 	double w = screenSize.getWidth()*800/1680;
				 	double h = screenSize.getHeight()*600/1050;
				 	Dimension d = new Dimension();
				 	d.setSize(w, h);
	            	
	            	checkPanel.setPreferredSize(d);
	            	JLabel note = new JLabel("               *Please select and download database from SHIPLYS server."+"\n");
	            	note.setFont(new Font("Arial", Font.BOLD,14));
	            	if(test_name.length>result_name.length){
	            	for (i=0;i<result_size;i++){
          		
	            		getCheckBox()[i] = new JCheckBox(i+1+". "+ test_name[i]);
	            		getCheckBox()[i].setName(i+1+". "+ test_name[i]);
	            		
	               		checkPanel.add(checkBox[i]);
	            	}}
	            	else{
	            		for (i=0;i<test_name.length;i++){
		            		
	            			getCheckBox()[i] = new JCheckBox(i+1+". "+ test_name[i]);
	            			getCheckBox()[i].setName(i+1+". "+ test_name[i]);
		            		
		            		checkPanel.add(checkBox[i]);
	            	}
	            		for (i=test_name.length;i<result_size;i++){
	            		
	            			getCheckBox()[i] = new JCheckBox(i+1+". empty");
	            			getCheckBox()[i].setName(i+1+". empty");
		            		
		               		checkPanel.add(checkBox[i]);
	            	}
	            	}
	            	DefaultTableModel DTM_list = new DefaultTableModel(data_list, columnNames);
	            	JTable table_odb = new JTable(DTM_list);
	            	JPanel panel_odb = new JPanel();
	            	
	            	JScrollPane sp_odb = new JScrollPane(table_odb);
					 panel_odb.setLayout(new BoxLayout(panel_odb,BoxLayout.Y_AXIS));
				     panel_odb.setName("online database");
				     
				     panel_odb.add(sp_odb);
				     
				     table_odb.setAlignmentX(Component.CENTER_ALIGNMENT);
				     
				 	//create a button for database download
				    button_download = new JButton("Download");
				    button_download.setName("Download");
				    button_download.setFont(font_1);
				 	
				    double w1 = screenSize.getWidth()*100/1680;
				 	double h1 = screenSize.getHeight()*20/1050;
				 	Dimension d1 = new Dimension();
				 	d1.setSize(w1, h1);
				    
				    button_download.setPreferredSize(d1);
				    button_download.addActionListener(createDownloadListener(0));

	 			     frame_odb.setLayout(new BorderLayout());
	 			     frame_odb.add(note,BorderLayout.PAGE_START);
	 			     frame_odb.add(checkScrollPanel,BorderLayout.CENTER);
	 			     frame_odb.add(button_download,BorderLayout.PAGE_END);
	 			     frame_odb.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
	 			     frame_odb.setResizable(true);
	 			     frame_odb.pack();
	 			     frame_odb.setVisible(true); 
	             }
		 };	
		 return searchListener;
		 }
		 
	//add actionlistner to button to download data from server to local database
		 private ActionListener createDownloadListener(int j){
			 ActionListener downloadLinster=  new ActionListener() {
	             @Override
	             public void actionPerformed(ActionEvent iL) {
	            	 for(i=0;i<result_size;i++){
	            			 if(getCheckBox()[i].isSelected()){
	            				 if(getCheckBox()[i].getName().endsWith("empty")){}
	            				 else{
	        	            		 System.out.println(getCheckBox()[i].getName());
	            				 }
	            	 }
	            	
	            	 
	            	 }
	             }
		 };	
		 return downloadLinster;
		 }	 
		 
		 

		 //main function
//main function to run to show GUI	 
	public static void main(String[] args) {
            

		
		if(args!=null && args.length>0) {
			
//to do a batch script run:
//"\Program Files (x86)\SHIPLYS LCT\SHIPLYS LCT.exe"  -i http://192.168.42.33:10000/data/v1/dbs/dc=shiplys.example -p "SHIPLYS Project"
		
			try {
				urlString = args[1];
				projectString = args[3];
				createAndShowGUI();
			} catch (BiffException | IOException e1) {
				JFrame F_warn0 = new JFrame("Error!");
				JPanel P_warn0 = new JPanel();
				P_warn0.setLayout(new BorderLayout());
				JLabel Fd_warn0 = new JLabel("Can't start up!");
				F_warn0.setSize(200, 80);
				P_warn0.add(Fd_warn0,BorderLayout.CENTER);
				F_warn0.setLocation(w_mid, h_mid);
				F_warn0.add(P_warn0);
				F_warn0.setVisible(true);
				F_warn0.setResizable(false); 	        					e1.printStackTrace();
			}
		}
		
		else {
		javax.swing.SwingUtilities.invokeLater(new Runnable() {
			
            public void run() {
            	JFrame F_id = new JFrame("New project");
            	F_id.setIconImage(new ImageIcon(cwd+"/pic/icon.png").getImage());
        		JPanel P_id = new JPanel();
        		
        	 	double w1 = screenSize.getWidth()*400/1680;
			 	double h1 = screenSize.getHeight()*120/1050;
			 	Dimension d1 = new Dimension();
			 	d1.setSize(w1, h1);
        		
        		P_id.setLayout(new BorderLayout());
        		JTextField TF = new JTextField();
        		
        	 	double w = screenSize.getWidth()*300/1680;
			 	double h = screenSize.getHeight()*45/1050;
			 	Dimension d = new Dimension();
			 	d.setSize(w, h);
        		
        		TF.setPreferredSize(d);
        		TF.setText("Enter project name here!"); 
        		JButton enter_name = new JButton();
        		enter_name.setText("Create a project");
        		enter_name.setFont(new Font("Arial", Font.BOLD,14));
        		TF.setBackground(Color.green);
        		enter_name.setBackground(Color.red);
        		enter_name.setForeground(Color.white);
        		enter_name.setPreferredSize(d);
        		ActionListener en = new ActionListener() {
        			
        			@Override
        			public void actionPerformed(ActionEvent e) {
        				frame_name = TF.getText();
        				try {
        					createAndShowGUI();
                                                
        				} catch (BiffException | IOException e1) {
							JFrame F_warn0 = new JFrame("Error!");
							JPanel P_warn0 = new JPanel();
							P_warn0.setLayout(new BorderLayout());
							JLabel Fd_warn0 = new JLabel("Can't start up!");
							F_warn0.setSize(200, 80);
							P_warn0.add(Fd_warn0,BorderLayout.CENTER);
							F_warn0.setLocation(w_mid, h_mid);
							F_warn0.add(P_warn0);
							F_warn0.setVisible(true);
							F_warn0.setResizable(false); 	        					e1.printStackTrace();
        				}
        				F_id.dispose();
        			}
        		};
        		enter_name.addActionListener(en);
        		TF.addActionListener(en);
        		F_id.setSize(d1);
//			 	P_id.setPreferredSize(d1);
        		P_id.add(TF,BorderLayout.PAGE_START);
        		P_id.add(enter_name,BorderLayout.PAGE_END);
        		F_id.add(P_id);
        		F_id.setVisible(true);
        		F_id.setResizable(false);  
        		F_id.setLocation(w_mid, h_mid);
        		F_id.setAlwaysOnTop(true);
        		F_id.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		
			}
		}
            
		);
		}
	}
	
//create frame
	private static void createAndShowGUI() throws BiffException, IOException {
  
	if(projectString == null) {
		frame = new JFrame("ShipLCA - "+ frame_name+" - " +dateFormat.format(date));
	}
	else {
		frame = new JFrame("ShipLCA - "+ projectString+" - " +dateFormat.format(date));
	}
		
		
	
	if(projectString ==null) {
	frame.setName(project_name);
	}
	else {
	frame.setName(projectString);
	}
	frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    
// 	double w = screenSize.getWidth()*1680/1680;
// 	double h = screenSize.getHeight()*1000/1050;
// 	Dimension d = new Dimension();
// 	d.setSize(w, h);
//    
//    frame.setPreferredSize(d);
    
    //Create and set up the content pane.
    JComponent newContentPane = new Gui16052019();
    frame.setLayout(new BorderLayout());
    frame.add(newContentPane,BorderLayout.WEST);
    frame.setExtendedState(JFrame.MAXIMIZED_BOTH);
    frame.setResizable(true);
    
 	double w1 = screenSize.getWidth()*1630/1680;
 	double h1 = screenSize.getHeight()*850/1050;
 	Dimension d1 = new Dimension();
 	d1.setSize(w1, h1);
        
    panel0.setPreferredSize(d1);//new Dimension(frame.getWidth()-35, frame.getHeight()-120)

    //Add menu
	//Create the menu bar.
	menuBar = new JMenuBar();
	
	//Build the first menu.
	menu = new JMenu("File");
	menu.setMnemonic(KeyEvent.VK_F);
	menu.getAccessibleContext().setAccessibleDescription(
	        "Help related actions");
	menuBar.add(menu);
	
	//a group of JMenuItems

	//to db	
	
		menuItem = new JMenuItem("Database ", KeyEvent.VK_D);
		menuItem.setAccelerator(KeyStroke.getKeyStroke(
		        KeyEvent.VK_D, ActionEvent.CTRL_MASK));
		menuItem.getAccessibleContext().setAccessibleDescription(
		        "Open database folder");
		AL_3=  new ActionListener() {

			public void actionPerformed(ActionEvent Menu) {
				Desktop desktop = Desktop.getDesktop();
				File folder = new File(cwd+"/db");
				
				    try {
						desktop.open(folder);;
					} catch (IOException e) {
						JFrame F_warn0 = new JFrame("Error!");
						JPanel P_warn0 = new JPanel();
						P_warn0.setLayout(new BorderLayout());
						JLabel Fd_warn0 = new JLabel("Can't open database folder! ");
						F_warn0.setSize(200, 80);
						P_warn0.add(Fd_warn0,BorderLayout.CENTER);
						F_warn0.setLocation(w_mid, h_mid);
						F_warn0.add(P_warn0);
						F_warn0.setVisible(true);
						F_warn0.setResizable(false); 						e.printStackTrace();
					}}};
		menuItem.addActionListener(AL_3);
		menu.add(menuItem);
	//to result
		menuItem = new JMenuItem("Reports ", KeyEvent.VK_R);
		menuItem.setAccelerator(KeyStroke.getKeyStroke(
		        KeyEvent.VK_R, ActionEvent.CTRL_MASK));
		menuItem.getAccessibleContext().setAccessibleDescription(
		        "Open report folder");
		AL_4=  new ActionListener() {

			public void actionPerformed(ActionEvent Menu) {
				Desktop desktop = Desktop.getDesktop();
				File folder = new File(cwd+"/reports");
				
				    try {
						desktop.open(folder);;
					} catch (IOException e) {
						JFrame F_warn0 = new JFrame("Error!");
						JPanel P_warn0 = new JPanel();
						P_warn0.setLayout(new BorderLayout());
						JLabel Fd_warn0 = new JLabel("Can't open report folder! ");
						F_warn0.setSize(200, 80);
						P_warn0.add(Fd_warn0,BorderLayout.CENTER);
						F_warn0.setLocation(w_mid, h_mid);
						F_warn0.add(P_warn0);
						F_warn0.setVisible(true);
						F_warn0.setResizable(false);						e.printStackTrace();
					}}};
		menuItem.addActionListener(AL_4);
		menu.add(menuItem);		
		
	//exit
		menuItem = new JMenuItem("Close", KeyEvent.VK_W);
		menuItem.setAccelerator(KeyStroke.getKeyStroke(
		        KeyEvent.VK_W, ActionEvent.CTRL_MASK));
		menuItem.getAccessibleContext().setAccessibleDescription(
		        "Exit");
		AL_4=  new ActionListener() {

			public void actionPerformed(ActionEvent Menu) {
				frame.dispatchEvent(new WindowEvent(frame, WindowEvent.WINDOW_CLOSING));
					}};
		menuItem.addActionListener(AL_4);
		menu.add(menuItem);	
		
	//Build the second menu.
	menu = new JMenu("Help");
	menu.setMnemonic(KeyEvent.VK_A);
	menu.getAccessibleContext().setAccessibleDescription(
	        "Help related actions");
	menuBar.add(menu);

	//a group of JMenuItems

//to website
	
	menuItem = new JMenuItem("About SHIPLYS", KeyEvent.VK_W);
	menuItem.setAccelerator(KeyStroke.getKeyStroke(
	        KeyEvent.VK_W, ActionEvent.CTRL_MASK));
	menuItem.getAccessibleContext().setAccessibleDescription(
	        "Open SHIPLYS website");
	AL_1=  new ActionListener() {

		public void actionPerformed(ActionEvent Menu) {
			Desktop desktop = Desktop.getDesktop();
			
			    try {
					desktop.browse(new URI("http://www.shiplys.com/"));
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (URISyntaxException e) {
					JFrame F_warn0 = new JFrame("Error!");
					JPanel P_warn0 = new JPanel();
					P_warn0.setLayout(new BorderLayout());
					JLabel Fd_warn0 = new JLabel("Can't open SHIPLYS website! ");
					F_warn0.setSize(200, 80);
					P_warn0.add(Fd_warn0,BorderLayout.CENTER);
					F_warn0.setLocation(w_mid, h_mid);
					F_warn0.add(P_warn0);
					F_warn0.setVisible(true);
					F_warn0.setResizable(false);					e.printStackTrace();
				}
			
			}};
	menuItem.addActionListener(AL_1);
	menu.add(menuItem);

//to manual

	menuItem = new JMenuItem("Documentation", KeyEvent.VK_D);
	menuItem.setAccelerator(KeyStroke.getKeyStroke(
	        KeyEvent.VK_D, ActionEvent.CTRL_MASK));
	menuItem.getAccessibleContext().setAccessibleDescription(
	        "Open Manual");
	AL_2=  new ActionListener() {

		public void actionPerformed(ActionEvent Menu) {
			Desktop desktop = Desktop.getDesktop();
			File manual = new File(cwd+"/Help.pdf");
			
			    try {
					desktop.open(manual);
				} catch (IOException e) {
					JFrame F_warn0 = new JFrame("Error!");
					JPanel P_warn0 = new JPanel();
					P_warn0.setLayout(new BorderLayout());
					JLabel Fd_warn0 = new JLabel("Can't open manual! ");
					F_warn0.setSize(200, 80);
					P_warn0.add(Fd_warn0,BorderLayout.CENTER);
					F_warn0.setLocation(w_mid, h_mid);
					F_warn0.add(P_warn0);
					F_warn0.setVisible(true);
					F_warn0.setResizable(false);					e.printStackTrace();
				}}};
	menuItem.addActionListener(AL_2);
	menu.add(menuItem);

//to dir	
	
	menuItem = new JMenuItem("Training Material ", KeyEvent.VK_T);
	menuItem.setAccelerator(KeyStroke.getKeyStroke(
	        KeyEvent.VK_T, ActionEvent.CTRL_MASK));
	menuItem.getAccessibleContext().setAccessibleDescription(
	        "Open Training Material");
	AL_3=  new ActionListener() {

		public void actionPerformed(ActionEvent Menu) {
			Desktop desktop = Desktop.getDesktop();
			File TM = new File(cwd+"/Training Material.pdf");
			
			    try {
					desktop.open(TM);;
				} catch (IOException e) {
					JFrame F_warn0 = new JFrame("Error!");
					JPanel P_warn0 = new JPanel();
					P_warn0.setLayout(new BorderLayout());
					JLabel Fd_warn0 = new JLabel("Can't open training material! ");
					F_warn0.setSize(200, 80);
					P_warn0.add(Fd_warn0,BorderLayout.CENTER);
					F_warn0.setLocation(w_mid, h_mid);
					F_warn0.add(P_warn0);
					F_warn0.setVisible(true);
					F_warn0.setResizable(false);					e.printStackTrace();
				}}};
	menuItem.addActionListener(AL_3);
	menu.add(menuItem);

	frame.setJMenuBar(menuBar);
	frame.setIconImage(new ImageIcon(cwd+"/pic/icon.png").getImage());
//    
    //Display the window.
    frame.pack();
    frame.setVisible(true);
    }
}