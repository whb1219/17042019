/*
 * SettingRetrievingCapex.java
 *
 * @date	 01.04.2019
 *
 * The contents of this file are subject to the terms of Atlantec-es' Development
 * and Distribution License (the License). You may not use this file except in
 * compliance with the License.
 *
 * You can obtain a copy of the License at http://www.atlantec-es.com/aes-ddl.html
 * or http://www.atlantec-es.com/aes-ddl.txt.
 *
 * Copyright (C) 1999-2019, Atlantec Enterprise Solutions GmbH
 */
package LCT;

import java.net.URI;
import java.net.URISyntaxException;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.Date;
import javax.ws.rs.core.UriBuilder;
import org.apache.log4j.Logger;
import org.atlantec.binding.erm.ActivityInstance;
import org.atlantec.db.Session;
import org.atlantec.directory.ConnectionMode;
import org.atlantec.directory.PooledPublisher;
import org.atlantec.objects.ObjectManager;
import static org.atlantec.rest.client.data.DataClient.DATA_SERVICE_PATH_PREFIX;
import static org.atlantec.rest.client.data.DataClient.DATA_SERVICE_VERSION;
import org.atlantec.rest.client.meta.MetaClient;
import static org.atlantec.rest.client.meta.MetaClient.META_SERVICE_PATH_PREFIX;
import static org.atlantec.rest.client.meta.MetaClient.META_SERVICE_VERSION;
import org.atlantec.rest.model.meta.IdefActivity;
import org.atlantec.rest.model.meta.ProcessModel;
import org.atlantec.util.DateHelper;
import org.atlantec.util.logging.LoggingSystem;
import org.jscience.economics.money.Currency;
import static org.jscience.economics.money.Currency.EUR;

/**
 *
 */
public class CreateActivityInstance
{
    static private final Logger log = Logger.getLogger(CreateActivityInstance.class);

    public static final int DATA_SERVICE_PORT = 10000;
    public static final int META_SERVICE_PORT = 10001;

    public static final URI DATA_SERVICE_URI;
    public static final URI META_SERVICE_URI;

    static {
        try {
            DATA_SERVICE_URI = new URI("http://localhost:" + DATA_SERVICE_PORT + DATA_SERVICE_PATH_PREFIX + DATA_SERVICE_VERSION);
            META_SERVICE_URI = new URI("http://localhost:" + META_SERVICE_PORT + META_SERVICE_PATH_PREFIX + META_SERVICE_VERSION);
        } catch (URISyntaxException ex) {
            throw new IllegalStateException("cannot initialize service URIs");
        }
    }

    private static String shipName;

    /**
     * create a new instance of CreateActivityInstance
     */
    public CreateActivityInstance()
    {
        super();
    }

    public static Session setUpDataServiceConnection(String connectionURL, String projectName) throws URISyntaxException
    {
        // init reference currency, otherwise working with cost-based measures will fail.
        Currency.setReferenceCurrency(EUR);
        if(connectionURL==null) {
        connectionURL = "http://localhost:10000/data/v1/dbs/dc=shiplys.example";
        }
        String username = System.getProperty("user.name", "tgadmin");
        char[] pw = "none".toCharArray();
//        String projectName = "SHIPLYS Project";
        if(projectName==null) {
            projectName = "SHIPLYS Project";
           }
        shipName = "MPC-100-Example";

//        URI connectionURI = UriBuilder.fromUri(DATA_SERVICE_URI).segment("dbs", "dc=shiplys.example").build();
        Session session = Session.newSession(new URI(connectionURL), username, pw, ConnectionMode.PROJECT_MODE, projectName, shipName);

        PooledPublisher.setApplicationDetails("Create Activity Instance",
                CreateActivityInstance.class, "1", CreateActivityInstance.class, "1");
        return session;
    }

    public static void main(String args[]) throws URISyntaxException, InterruptedException
    {
        LoggingSystem.setUp();
        //Logger.getRootLogger().setLevel(Level.DEBUG);
        Session session = setUpDataServiceConnection("", "");

        CreateActivityInstance createActivityInstance = new CreateActivityInstance();

        createActivityInstance.create(session.getManager(),"A124");
        session.close();
    }

    public ActivityInstance create(ObjectManager projOM,String activityName)
    {
        //retrieve the activity which is referenced by the activity instance
        MetaClient metaClient = MetaClient.getInstance(META_SERVICE_URI, null);
        ProcessModel shiplysModel = metaClient.getProcessModel("SHIPLYS");

        IdefActivity activity = shiplysModel.getActivity(activityName);

        //create new activity instance
        ActivityInstance actInst = projOM.createInformationObject(ActivityInstance.class, activity
                .getPathName()+".instance");

        //set the plan attribute of activity instance referencing the activity received from meta service
        actInst.setPlan(activity);

        //TODO: start date could be set by the work flow controller after selecting a specific activity,
        //the appropriate software and starting a job
        Date startDate = Date.from(Instant.now().minus(5, ChronoUnit.MINUTES));
        actInst.setStartDate(DateHelper.formatIsoDate(startDate));

        Date endDate = Date.from(Instant.now());
        actInst.setEndDate(DateHelper.formatIsoDate(endDate));

        projOM.makePersistent(actInst);
        projOM.currentTransaction().commit();

        return actInst;
    }

}   // class CreateActivityInstance //
