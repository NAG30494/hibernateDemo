/**
 * 
 */
package org.hibernate.client;

import org.hibernate.HibernateException;
import org.hibernate.Session;
import org.hibernate.Transaction;
import org.hibernate.domain.Message;
import org.hibernate.utils.HibernateUtils;

/**
 * @author Bhaviya
 *
 */
public class MessageClient {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		System.out.println("test!!!!");
		System.out.println("Hello!");
		Session session = null;
		Transaction transaction = null;
		Message message = null;
		try {
			message = new Message("Hello Hibernate!!!");
			session = HibernateUtils.getSessionFactory().openSession();
			transaction = session.getTransaction();
			transaction.begin();
			session.save(message);
			transaction.commit();
		} catch (HibernateException e) {
			e.printStackTrace();
		} finally {
			System.out.println("inside finally");
			if (session != null) {
				session.close();
				session = null;
			}

		}

	}

}




