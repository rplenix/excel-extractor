import arrow.core.Either
import java.io.FileNotFoundException
import java.io.InputStream

fun openResource(resourceName: String): Either<FileNotFoundException, InputStream> =
    Either.fromNullable(Thread.currentThread().contextClassLoader.getResourceAsStream(resourceName))
        .mapLeft { FileNotFoundException("No such resource: $resourceName") }