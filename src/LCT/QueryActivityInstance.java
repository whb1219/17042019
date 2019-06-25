/*
 * QueryExamplesGeneral.java
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

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Set;
import static java.util.stream.Collectors.joining;
import javax.ws.rs.core.UriBuilder;
import static org.apache.log4j.Level.DEBUG;
import static org.apache.log4j.Level.INFO;
import org.apache.log4j.Logger;
import org.atlantec.binding.POID;
import org.atlantec.binding.erm.ActivityInstance;
import org.atlantec.binding.erm.ActivityTemplate;
import org.atlantec.db.Session;
import org.atlantec.directory.ConnectionMode;
import org.atlantec.directory.PooledPublisher;
import org.atlantec.directory.event.Event;
import org.atlantec.directory.event.EventCategory;
import org.atlantec.directory.event.State;
import org.atlantec.directory.event.changeset.TGChangeSetState;
import org.atlantec.directory.schema.EventLog;
import org.atlantec.objects.FilterExpression;
import org.atlantec.objects.ObjectManager;
import org.atlantec.objects.Query;
import org.atlantec.rest.base.mapping.ResourceDeserializer;
import static org.atlantec.rest.client.data.DataClient.DATA_SERVICE_PATH_PREFIX;
import static org.atlantec.rest.client.data.DataClient.DATA_SERVICE_VERSION;
import org.atlantec.rest.client.meta.MetaClient;
import static org.atlantec.rest.client.meta.MetaClient.META_SERVICE_PATH_PREFIX;
import static org.atlantec.rest.client.meta.MetaClient.META_SERVICE_VERSION;
import org.atlantec.rest.model.meta.IdefActivity;
import org.atlantec.rest.model.meta.ProcessModel;
import org.atlantec.util.CollectionsHelper;
import static org.atlantec.util.StringHelper.LINEFEED;
import org.atlantec.util.logging.LoggingSystem;
import static org.atlantec.util.logical.LogicalOperator.EXACT;
import org.jscience.economics.money.Currency;
import static org.jscience.economics.money.Currency.EUR;

/**
 * To execute a Query the following steps need to be done: - Start Session - Init ObjectManager - Create query - Set search mode
 * - Set search filter - Execute query - Do processing with query result - Close session
 * <p>
 * Queries can be executed in three different modes: - Persistent: Search in the data base Search does not include uncommitted
 * local changes Search filter can only be set on indexed attributes (f.e. POID, Parent, Type, TypeX, CommonName, ...) - Local:
 * Searches the local application memory for objects Setting filters on every attribute (f.e. Length, weight) of the object is
 * possible Loading clustered objects Does not include changes in the data base probably committed by another user at the same
 * time - Full: Persistent and local search If there are two different versions of an InformationObject in the data base and in
 * application memory the latest version of the object is returned
 */
public class QueryActivityInstance
{
    static private final Logger log = Logger.getLogger(QueryActivityInstance.class);

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

    private static ObjectManager projOM;
    private static EventLog tgcssEventLog;
    private static State committedState;

    /**
     * create a new instance of QueryExamplesGeneral
     */
    public QueryActivityInstance()
    {
        super();
    }

    public static void setUpDataServiceConnection() throws URISyntaxException
    {
        // init reference currency, otherwise working with cost-based measures will fail.
        Currency.setReferenceCurrency(EUR);
        String username = System.getProperty("user.name", "tgadmin");
        char[] pw = "none".toCharArray();
        String projectName = "SHIPLYS Project";
        String shipName = "MPC-100-Example";

        //session related to project mode
        URI connectionURI = UriBuilder.fromUri(DATA_SERVICE_URI).segment("dbs", "dc=shiplys.example").build();
        Session session = Session.newSession(connectionURI, username, pw, ConnectionMode.PROJECT_MODE, projectName, shipName);

        projOM = session.getManager();
        EventCategory tgChangeSetStateEC = projOM.getContext().getSystemEventCategory(TGChangeSetState.class);
        tgcssEventLog = projOM.getContext().getEventLog(tgChangeSetStateEC);
        committedState = tgChangeSetStateEC.getStateDefinition(TGChangeSetState.committed.name()).createState();

        PooledPublisher.setApplicationDetails("Design Management Tool",
                QueryActivityInstance.class, "1", QueryActivityInstance.class, "1");
    }

    public static void main(String args[]) throws URISyntaxException, InterruptedException
    {
        LoggingSystem.setUp();
        //Logger.getRootLogger().setLevel(Level.DEBUG);

        setUpDataServiceConnection();

        QueryActivityInstance queryActivityInstance = new QueryActivityInstance();

       // Logger.getLogger(ResourceDeserializer.class).setLevel(DEBUG);
       // Logger.getLogger(AbstractClient.class).setLevel(DEBUG);
        queryActivityInstance.retrieveActivityInstance();
    }

    public ActivityInstance retrieveActivityInstance()
    {
        log.info("******persistentQueryWithSimpleFilter");

        MetaClient metaClient = MetaClient.getInstance(META_SERVICE_URI, null);
        ProcessModel shiplysModel = metaClient.getProcessModel("SHIPLYS");
        
        FilterExpression filter = new FilterExpression(ActivityInstance.COMMONNAME, EXACT, "A127.instance");

        Logger.getLogger(ResourceDeserializer.class).setLevel(DEBUG);
        Set<ActivityInstance> result = projOM.queryAndLoad(ActivityInstance.class, Query.Mode.PERSISTENT, filter);

        ActivityInstance actInstance;
        if (CollectionsHelper.isNullOrEmpty(result)) {
            throw new UnsupportedOperationException("There is no activity with the name A127!");
        } else if (result.size() > 1) {
            throw new UnsupportedOperationException("There are several activities having the same name A127!");
        } else {
            actInstance = CollectionsHelper.getFirstItem(result);
        }
        ActivityTemplate plan = actInstance.getPlan();
        if (plan instanceof IdefActivity) {
            IdefActivity act = (IdefActivity)plan;
          //TODO: this kind of resolution does not work yet
          //  List<Parameter> controls = act.getControls();
        }
        Logger.getLogger(ResourceDeserializer.class).setLevel(INFO);
        IdefActivity activity = shiplysModel.getActivity(plan.getPOID().getName());
        log.info("Activity " + activity.getCommonName() + " " + activity.getName());

        log.info("Activity " + actInstance.getPathName() + " (" + activity.getName() + ": " + activity.getDescription()
                + ") was completed at " + actInstance.getEndDate());
        Set<Event> events = tgcssEventLog.getEvents(actInstance.getPOID(), null, committedState);
        log.info("A127: TGChangeSetState.committed events: " + CollectionsHelper.formatAsColumn(events));
        if (CollectionsHelper.isNotEmpty(events)) {
            Event ev = events.iterator().next();
            log.info("Event " + ev.getPOID() + ": source=" + ev.getSource() + ", targets= " + LINEFEED
                    + ev.getTargets().stream()
                            .map(t -> new POID(t))
                            .map(p -> p.getType().getSimpleName() + " " + p.getName())
                            .collect(joining(LINEFEED))
            );
        }
        log.debug(this);

        return actInstance;
    }

}   // class QueryExamplesGeneral //

