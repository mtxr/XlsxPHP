namespace XlsxPhp;

use XlsxPhp\Writer;

class BufferedWriter
{
    protected fd         = null;
    protected buffer     = "";
    protected check_utf8 = false;
    protected bufferSize = 8191;

    public function __construct(filename, fd_fopen_flags = "w", check_utf8=false, bufferSize = 8191)
    {
        let this->check_utf8 = check_utf8;
        let this->bufferSize = bufferSize;
        let this->fd = fopen(filename, fd_fopen_flags);
        if (this->fd===false) {
            Writer::log("Unable to open " . filename . " for writing.");
        }
    }

    public function write(message)
    {
        let this->buffer = this->buffer . message;
        if (strlen(this->buffer) > this->bufferSize) {
            return this->purge();
        }
        return this;
    }

    protected function purge()
    {
        if (this->fd) {
            if (this->check_utf8 && !self::isValidUTF8(this->buffer)) {
                Writer::log("Error, invalid UTF8 encoding detected.");
                let this->check_utf8 = false;
            }
            fwrite(this->fd, this->buffer);
            let this->buffer = "";
        }
        return this;
    }

    public function close()
    {
        this->purge();
        if (this->fd) {
            fclose(this->fd);
            let this->fd = null;
        }
    }

    public function __destruct()
    {
        this->close();
    }

    public function ftell()
    {
        if (this->fd) {
            this->purge();
            return ftell(this->fd);
        }
        return -1;
    }

    public function fseek(pos)
    {
        if (this->fd) {
            this->purge();
            return fseek(this->fd, pos);
        }
        return -1;
    }

    protected static function isValidUTF8(message)
    {
        if (function_exists("mb_check_encoding"))
        {
            return mb_check_encoding(message, "UTF-8") ? true : false;
        }
        return preg_match("//u", message) ? true : false;
    }
}

// vim: set filetype=php expandtab tabstop=4 shiftwidth=4 autoindent smartindent:
