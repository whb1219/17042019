/*
 * GUIexample.java
 * 
 * @date	 30.10.2018
 * 
 * The contents of this file are subject to the terms of Atlantec-es' Development
 * and Distribution License (the License). You may not use this file except in
 * compliance with the License.
 * 
 * You can obtain a copy of the License at http://www.atlantec-es.com/aes-ddl.html
 * or http://www.atlantec-es.com/aes-ddl.txt.
 * 
 * Copyright (C) 1999-2018, Atlantec Enterprise Solutions GmbH
 */
package LCT;


import java.util.Set;
import org.apache.log4j.Logger;
import org.atlantec.binding.erm.CatalogueItem;
import org.atlantec.binding.erm.KeyValue;
import org.atlantec.binding.erm.ProductComponent;
import org.atlantec.util.KeyValueHelper;
import org.atlantec.util.logging.LoggingSystem;

/**
 *
 */
public class GUIexample
{
    static private final Logger log = Logger.getLogger(GUIexample.class);

    public static void main(String[] args) throws Exception
    {
        LoggingSystem.setUp();
        //Logger.getRootLogger().setLevel(Level.DEBUG);

        //setup data service connection
        LCAdataFactory.setUpDataServiceConnection(null, null);

        
        //a common cost classification has been already created and is available within data base, that's why the method is outcommented
        //LCAdataFactory.createCommonCostClassification();

        //create a new catalogue
        //Note: it is not possible to create a catalogue if already one exist having the same name
        //createCatalogue methode creates a completly new catalogue and since this method has been already executed the related
        //catalogue already exists in the database, that's why the method is outcommented
        //LCAdataFactory.createCatalogue();


        GUIexample guiExample = new GUIexample();

        //Shows the default property values of the engine CatalogueItem
        guiExample.showDefaultEngineValues();

        //Shows the default property values of the battery CatalogueItem
        guiExample.showDefaultBatteryValues();


       //Note: already exists in the database, that's why the method is outcommented
//        //Create new engine product component
//        ProductComponent engine = guiExample.createEngine("EngineXYZ");
//
//        //Replace the default price and mass value of the new engine
//        String priceValue = new MoneyMeasure(267000, Currency.EUR).toString();
//        LCAdataFactory.setProductComponentsProperty(engine, "price", priceValue);
//
//        String massValue = new MassMeasure(777, NonSI.TON_UK).toString();
//        LCAdataFactory.setProductComponentsProperty(engine, "mass", massValue);
//
//        //Create new battery product component
//        ProductComponent battery = guiExample.createBattery("BatteryXYZ");
//
//        //Create Scenario
//        AnalysisScenario scenario = LCAdataFactory.createAnalysisScenario("lcaAnalysisScenario");
//
//        //Create Case
//        AnalysisCase analysisCase = LCAdataFactory.createAnalysisCase(scenario, "lcaAnalysisCase");
//
//        //Create General LCA Parameters
//        ParameterSet lcaGeneralParameterSet = LCAdataFactory.createParameterSet(analysisCase, "GeneralParameters",
//                "This set contains the common lca parameters.", "SHIPLYS general LCA parameter definitions");
//
//        LCAdataFactory.setLifeSpan(lcaGeneralParameterSet, 25);
//        LCAdataFactory.setPresentValue(lcaGeneralParameterSet, 75000000);
//        LCAdataFactory.setInterestRate(lcaGeneralParameterSet, 5);
//        LCAdataFactory.setSensitivityLevel(lcaGeneralParameterSet, 0);
//        LCAdataFactory.setShipTotalPrice(lcaGeneralParameterSet, 100000000);
//
//        //---log general LCA parameters of the analysis case
//        log.info("******General");
//        for (KeyValue kv: lcaGeneralParameterSet.getParameters()) {
//            log.info("Parameter name =" + kv.getKey() + ", Parameter type =" + kv.getTypeX() + ", Parameter value =" + kv
//                    .getValue() + "\n");
//        }
//
//        //Create Construction LCA Parameters
//        //TODO: in the same way as for General LCA Parameters. The related setter methods have also to be created (see LCAdataCreation for which parameters)
//        //Create Operation LCA Parameters
//        //TODO: in the same way as for General LCA Parameters. The related setter methods have also to be created (see LCAdataCreation for which parameters)
//        //Create Scrapping LCA Parameters
//        //TODO: in the same way as for General LCA Parameters. The related setter methods have also to be created (see LCAdataCreation for which parameters)
//        //Create Evaluation Results
//        //TODO: in the same way as for General LCA Parameters. The related setter methods have also to be created (see LCAdataCreation for which parameters)
//        //Create Result
//        EvaluationResult result = LCAdataFactory.createEvaluationResult(analysisCase, "evaluationResult");
//
//        LCAdataFactory.setLifeCycleCost(result, 1500000);
//        //TODO: create in the same way as for LifeCycleCost setter method for other values (see LCAdataCreation createEvaluationResults method)
//
//        //Save
//        LCAdataFactory.projOM.makePersistent(scenario);
//        LCAdataFactory.projOM.makePersistent(analysisCase);
//        LCAdataFactory.projOM.makePersistent(result);
//
//        LCAdataFactory.projOM.currentTransaction().commit();

    }

    //For example if the user selects "Engine" from your list (in your GUI component) and the source is the data service instead of the local excel sheet,
    //then the related catalogue item is requested and the default values are shown.
    public Set<KeyValue> showDefaultEngineValues()
    {
        Set<CatalogueItem> engineItems = LCAdataFactory.getCatalogueItemsByClassName(LCAdataFactory.catalogueName, "ENGINE");

        //Note: we assume currently that only one engine catalogue item exists within the catalogue. If several catalogue items for different engine types
        //should exist in the future, then it should be possible for the user to select from different engine types containing default values and manipulate these if required.
        if (engineItems.size() > 1) {
            throw new UnsupportedOperationException(
                    "Several catalogue items representing different engine types containing default velaues are not supported yet!");
        }

        CatalogueItem engineItem = null;
        if (engineItems.iterator().hasNext()) {
            engineItem = engineItems.iterator().next();
        } else {
            throw new IllegalArgumentException("No catalogue item of a class ENGINE is available in catalogue!");
        }

        Set<KeyValue> engineItemProps = engineItem.getProperties();
        //Log the properties
        log.info("Properties of the Engine Item");
        for (KeyValue engingeItemProp: engineItemProps) {
            log.info("Parameter name =" + engingeItemProp.getKey() + ", Parameter type =" + engingeItemProp.getTypeX()
                    + ", Parameter value =" + KeyValueHelper.getTypedValue(engingeItemProp) + "\n");
        }

        return engineItemProps;
    }

    //For example if the user selects "Battery" from your list (in your GUI component) and the source is the data service instead of the local excel sheet,
    //then the related catalogue item is requested and the default values are shown
    public Set<KeyValue> showDefaultBatteryValues()
    {
        Set<CatalogueItem> batteryItems = LCAdataFactory.getCatalogueItemsByClassName(LCAdataFactory.catalogueName, "BATTERY");

        //Note: we assume currently that only one battery catalogue item exists within the catalogue. If several catalogue items for different engine types
        //should exist in the future, then it should be possible for the user to select from different engine types containing defualt values and manipulate these if required.
        if (batteryItems.size() > 1) {
            throw new UnsupportedOperationException(
                    "Several catalogue items representing different engine types containing default velaues are not supported yet!");
        }

        CatalogueItem batteryItem = null;
        if (batteryItems.iterator().hasNext()) {
            batteryItem = batteryItems.iterator().next();
        } else {
            throw new IllegalArgumentException("No catalogue item of a class BATTERY is available in catalogue!");
        }

        Set<KeyValue> batteryItemProps = batteryItem.getProperties();
        //Log the properties
        log.info("Properties of the Engine Item");
        for (KeyValue engingeItemProp: batteryItemProps) {
            log.info("Parameter name =" + engingeItemProp.getKey() + ", Parameter type =" + engingeItemProp.getTypeX()
                    + ", Parameter value =" + KeyValueHelper.getTypedValue(engingeItemProp) + "\n");
        }

        return batteryItemProps;
    }

    //Create new product component for engine by refering the related catalogueItem of the shipyard catalogue
    public ProductComponent createEngine(String engineName)
    {
        Set<CatalogueItem> engineItems = LCAdataFactory.getCatalogueItemsByClassName(LCAdataFactory.catalogueName, "ENGINE");

        //Note: we assume currently that only one engine item exist within the catalogue. If several catalogue items for different engine types
        //should exist in the future, then it should be possible for the user to select from different engine types containing defualt values and manipulate these if required.
        if (engineItems.size() > 1) {
            throw new UnsupportedOperationException(
                    "Several catalogue items representing different engine types containing default velaues are not supported yet!");
        }

        ProductComponent engine = null;
        if (engineItems.iterator().hasNext()) {
            CatalogueItem engineItem = engineItems.iterator().next();

            //Create a new engine having the example name "EngineXYZ"
            engine = LCAdataFactory.createProductComponent(engineItem, engineName);
        }

        LCAdataFactory.projOM.makePersistent(engine);

        return engine;
    }

    //Create new product component for battery by refering the related catalogueItem of the shipyard catalogue
    public ProductComponent createBattery(String batteryName)
    {
        Set<CatalogueItem> batteryItems = LCAdataFactory.getCatalogueItemsByClassName(LCAdataFactory.catalogueName, "BATTERY");

        //Note: we assume currently that only one engine item exist within the catalogue. If several catalogue items for different engine types
        //should exist in the future, then it should be possible for the user to select from different engine types containing defualt values and manipulate these if required.
        if (batteryItems.size() > 1) {
            throw new UnsupportedOperationException(
                    "Several catalogue items representing different battery types containing default velaues are not supported yet!");
        }

        ProductComponent battery = null;
        if (batteryItems.iterator().hasNext()) {
            CatalogueItem batteryItem = batteryItems.iterator().next();

            //Create a new engine having the example name "BatteryXYZ"
            battery = LCAdataFactory.createProductComponent(batteryItem, batteryName);
        }

        LCAdataFactory.projOM.makePersistent(battery);

        return battery;
    }

}   // class GUIexample //
