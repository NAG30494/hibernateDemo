package org.hibernate.client;

import java.util.ArrayList;
import java.util.List;

import org.hibernate.domain.Book;
import org.hibernate.domain.Chapter;
import org.hibernate.domain.Publisher;
import org.hibernate.service.BookStoreService;



public class BookStoreClient {
	public static void main(String[] args) {
		BookStoreService bookStoreService = new BookStoreService();

		//persisting object graph
		
		Publisher publisher = new Publisher("MANN", "Manning Publications Co.");

		Book book = new Book("9781617290459", "Java Persistence with Hibernate, Second Edition", publisher);

		List<Chapter> chapters = new ArrayList<Chapter>();
		Chapter chapter1 = new Chapter("Introducing JPA and Hibernate", 1);
		chapters.add(chapter1);
		Chapter chapter2 = new Chapter("Domain Models and Metadata", 2);
		chapters.add(chapter2);

		book.setChapters(chapters);

		bookStoreService.persistObjectGraph(book);
		
		//retrieving object graph
		
		book = bookStoreService.retrieveObjectGraph("9781617290459");
		System.out.println(book);
		

	}
}





