--postgressql

CREATE TYPE gender_type AS ENUM ('male', 'female')

CREATE TABLE users (
    usernmae CHAR(10) NOT NULL UNIQUE,
    email VARCHAR(50) NOT NULL UNIQUE,
    gender gender_type NOT NULL,
    interests TEXT[] NOT NULL,  -- SET in MySQL
    bio TEXT, -- No TINYTEXT, LONGTEXT etc exists in PostgreSQL
        --TOAST - The Oversied-Attribute Storage Technique
    profile_photo BYTEA, -- No BLOB exists in PostgreSQL
    cover_photo BYTEA, -- 
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    is_admin BOOLEAN NOT NULL DEFAULT FALSE,
    joined_at TIMESTAMP WITH TIMESTAMP TIME ZONE
)


CREATE TABLE genres (
    genre_id BIGINT PRIMARY KEY GENERATED ALWAYS AS IDENTITY,
    name VARCHAR(50) UNIQUE,
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP NOT NULL,

)

INSERT INTO
    genres (name)
SELECT DISTINCT UNNEST(STRING_TO_ARRAY(genres, ',')) AS genre 
FROM movies GROUP BY genres;


CREATE TABLE movies_genres (
    movie_id BIGINT NOT NULL,
    genre_id BIGINT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP NOT NULL,
    PRIMARY KEY (movie_id, genre_id),
    FOREIGN KEY (movie_id) REFERENCES movies(movie_Id),
    FOREIGN KEY (genre_id) REFERENCES genres(genre_id),
    
)

SELECT movies.title, movies.movie_id, genres.name, genres.genre_id
FROM movies
JOIN genres ON movies.genres LIKE '%' || genres.name || '%';


INSERT INTO movies_genres (movie_id, genre_id)
SELECT  movies.movie_id, genres.genre_id
FROM movies
JOIN genres ON movies.genres LIKE '%' || genres.name || '%';