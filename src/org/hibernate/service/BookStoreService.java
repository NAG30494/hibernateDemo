package org.hibernate.service;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.hibernate.domain.Book;
import org.hibernate.domain.Chapter;
import org.hibernate.domain.Publisher;

public class BookStoreService {
	private Connection connection = null;
	static String userName = "root";
	static String password = "admin";
	static String database = "hibernatedemo";

	public void initConnection() throws ClassNotFoundException,SQLException {
		Class.forName("com.mysql.jdbc.Driver");
		connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/"+database, userName, password);
	}
	
	private void closeConnection(){
		try {
			connection.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}

	public void persistObjectGraph(Book book) {
		try {
			initConnection();
			PreparedStatement stmt = connection.prepareStatement("insert into PUBLISHER (CODE, PUBLISHER_NAME) values (?, ?)");
			stmt.setString(1, book.getPublisher().getCode());
			stmt.setString(2, book.getPublisher().getName());
			stmt.executeUpdate();
			stmt.close();

			stmt = connection.prepareStatement("insert into BOOK (ISBN, BOOK_NAME, PUBLISHER_CODE) values (?, ?, ?)");
			stmt.setString(1, book.getIsbn());
			stmt.setString(2, book.getName());
			stmt.setString(3, book.getPublisher().getCode());
			stmt.executeUpdate();

			stmt.close();

			stmt = connection.prepareStatement("insert into CHAPTER (BOOK_ISBN, CHAPTER_NUM, TITLE) values (?, ?, ?)");
			for (Chapter chapter : book.getChapters()) {
				stmt.setString(1, book.getIsbn());
				stmt.setInt(2, chapter.getChapterNumber());
				stmt.setString(3, chapter.getTitle());
				stmt.executeUpdate();
			}
			stmt.close();
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			closeConnection();
		}
	}

	public Book retrieveObjectGraph(String isbn) {
		Book book = null;
		try {
			initConnection();
			PreparedStatement stmt = connection.prepareStatement("select * from BOOK, PUBLISHER where BOOK.PUBLISHER_CODE = PUBLISHER.CODE and BOOK.ISBN = ?");
			stmt.setString(1, isbn);

			ResultSet rs = stmt.executeQuery();

			book = new Book();
			if (rs.next()) {
				book.setIsbn(rs.getString("ISBN"));
				book.setName(rs.getString("BOOK_NAME"));

				Publisher publisher = new Publisher();
				publisher.setCode(rs.getString("CODE"));
				publisher.setName(rs.getString("PUBLISHER_NAME"));
				book.setPublisher(publisher);
			}

			rs.close();
			stmt.close();

			List<Chapter> chapters = new ArrayList<Chapter>();
			stmt = connection.prepareStatement("select * from CHAPTER where BOOK_ISBN = ?");
			stmt.setString(1, isbn);
			rs = stmt.executeQuery();

			while (rs.next()) {
				Chapter chapter = new Chapter();
				chapter.setTitle(rs.getString("TITLE"));
				chapter.setChapterNumber(rs.getInt("CHAPTER_NUM"));
				chapters.add(chapter);
			}
			book.setChapters(chapters);

			rs.close();
			stmt.close();
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			closeConnection();
		}
		return book;
	}

}
