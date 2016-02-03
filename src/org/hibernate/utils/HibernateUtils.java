package org.hibernate.utils;

import org.hibernate.SessionFactory;
import org.hibernate.boot.registry.StandardServiceRegistryBuilder;
import org.hibernate.cfg.Configuration;

public class HibernateUtils {
	
	private static SessionFactory sessionFactory = null;
	
    private static void buildSessionFactory() {
        try {        	
        	// Create the SessionFactory from hibernate.cfg.xml
        	Configuration configuration = new Configuration().configure("resources/hibernate.cfg.xml");    
        	System.out.println("Configuration completed!");
        	sessionFactory = configuration.buildSessionFactory();
           // return configuration.buildSessionFactory( new StandardServiceRegistryBuilder().applySettings( configuration.getProperties() ).build() );
        }
        catch (Throwable ex) {
            // Make sure you log the exception, as it might be swallowed
            System.err.println("Initial SessionFactory creation failed." + ex);
            throw new ExceptionInInitializerError(ex);
        }
    }

    public static SessionFactory getSessionFactory() {
    	if(sessionFactory==null){
    		buildSessionFactory();
    	}
        return sessionFactory;
    } 

}
