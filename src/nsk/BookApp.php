<?php
/**
 * Created by PhpStorm.
 * User: tobias
 * Date: 2017-08-31
 * Time: 19:31
 */

namespace nsk;


use Sabre\Xml\Service;

class BookApp
{

    /**
     *
     */
    public function run() {
        $service = new Service();
        $service->mapValueObject('{http://example.org/books}books', Books::class);
        $service->mapValueObject('{http://example.org/books}book', Book::class);

        $books = new Books();
        $books->book[] = new Book("Maja", "Widmark");
        $books->book[] = new Book("Lasse", "Kalle");
        $xml = $service->writeValueObject($books);
        echo $xml;
        echo "\n";
    }
}
class Books {

    // A list of books.
    public $book = [];

}
class Book {

        public function __construct($title, $author) {
            $this->author = $author;
            $this->title = $title;
        }
        public $title;
        public $author;

}