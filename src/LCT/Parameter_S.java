package LCT;

public class Parameter_S {
//Properties
	String 	Transportation_fuel_type;
//	String 	Electricity_type;
	String 	Fuel_type;
//	String 	LO_type;
	String 	Cut_type;
	
	double	Machinery_Number	=0;
	double	Machinery_Weight	=0;
	double	Hull_Weight=0;
	double	S_Price_M	=0;
	double	S_Price_H	=0;
	double	Transportation_distance	=0;
	double	Transportation_fee	=0;
	double	Transportation_SFOC	=0;
	double	Transportation_fuel_price	=0;
	double	Electricity_price	=0;
	
	double 	Fuel_price  =0;
	double 	Cutting_power=0;
	double  Cutting_hours=0;
	double 	Cleaning_power=0;
	double  Cleaning_hours=0;
	double  Cutting_material=0;
	double 	Cutting_M_price=0;
	double	Life_span	=0;
	double	Interest	=0;
	double 	PV=0;
	double  Cost_S=0;
	
	double Spec_GWP_Trans=0;
	double Spec_AP_Trans=0;
	double Spec_EP_Trans=0;
	double Spec_POCP_Trans=0;
	double Spec_GWP_E=0;
	double Spec_AP_E=0;
	double Spec_EP_E=0;
	double Spec_POCP_E=0;
	double Spec_GWP_FO=0;
	double Spec_AP_FO=0;
	double Spec_EP_FO=0;
	double Spec_POCP_FO=0;
//	double Spec_GWP_LO;
//	double Spec_AP_LO;
//	double Spec_EP_LO;
//	double Spec_POCP_LO;

	
	double GWP	=0;
	double AP 	=0;
	double EP	=0;
	double POCP	=0;
	
	public void run(){	
		//Interim results
		double Transportation_fuel_quantity=Transportation_SFOC*Transportation_distance/100*Machinery_Weight*Machinery_Number*1000/1000;
		double Transportation_fuel_cost= Transportation_fuel_price*Transportation_fuel_quantity;
		
		if(Cutting_hours==0) {
			Cutting_hours=Hull_Weight*20;
		}
		else {
			}
		
		double Electricity_consumption= Cutting_hours*Cutting_power+Cleaning_hours*Cleaning_power;
		
		double Electricity_cost = Electricity_consumption*Electricity_price;
		double Cutting_M_Cost = Cutting_material*Cutting_M_price;
		double S_Cost_hull = Hull_Weight*S_Price_H;
		double S_Cost_mach= Machinery_Number*S_Price_M;
		
		//Final result	
		//LCA
		GWP	= (Spec_GWP_Trans * Machinery_Number*Machinery_Weight*Transportation_distance/100 +Spec_GWP_E*Electricity_consumption/1000);
		AP 	= (Spec_AP_Trans * Machinery_Number*Machinery_Weight*Transportation_distance/100 +Spec_AP_E*Electricity_consumption/1000);
		EP	= (Spec_EP_Trans * Machinery_Number*Machinery_Weight*Transportation_distance/100 +Spec_EP_E*Electricity_consumption/1000);
		POCP= (Spec_POCP_Trans * Machinery_Number*Machinery_Weight*Transportation_distance/100 +Spec_POCP_E*Electricity_consumption/1000);		
				
		if (PV==0){
			Cost_S = Electricity_cost+S_Cost_hull+Cutting_M_Cost+Transportation_fuel_cost+S_Cost_mach;//Electricity_cost+
			}
		else{
			Cost_S = (Electricity_cost+S_Cost_hull+Cutting_M_Cost+Transportation_fuel_cost+S_Cost_mach)/(Math.pow((1+Interest),Life_span));//
			};
//			System.out.println("Scrapping: cost is : " + Cost_S +"Euro");
//			System.out.println("Scrapping: GWP is :" + GWP + "ton CO2e"); 
//			System.out.println("Scrapping: AP is :" + AP + "ton SO2e"); 
//			System.out.println("Scrapping: EP is :" + EP + "ton PO4e"); 
//			System.out.println("Scrapping: POCP is :" + POCP + "ton C2H6e"); 		
	}
}