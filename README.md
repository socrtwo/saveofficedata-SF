# Corrupt Office Data/Text Extract Service

<!--PAGES_LINK_BANNER-->
> 🌐 **Live page:** [https://socrtwo.github.io/saveofficedata-SF/](https://socrtwo.github.io/saveofficedata-SF/)  
> 📦 **Releases:** [github.com/socrtwo/saveofficedata-SF/releases](https://github.com/socrtwo/saveofficedata-SF/releases)
<!--/PAGES_LINK_BANNER-->

A web-based service that lets users upload corrupt Office files and extracts whatever text can be salvaged. Provides a simple upload-and-download interface.

## Screenshots

Visit the [SourceForge project page](https://sourceforge.net/projects/saveofficedata/) to view screenshots.

> **Tip:** If you have screenshots to contribute, open a PR adding them to a `screenshots/` folder!

**Language:** PHP  
**License:** MIT

## Features

- Web-based upload interface for corrupt Office files
- Extracts text from damaged DOCX, XLSX, PPTX files
- No software installation needed — works in any browser
- Download recovered text as a plain text file

## System Requirements

- PHP 7.0 or later
- A web server (Apache, Nginx, or PHP built-in server)
- MySQL/MariaDB (if the project uses a database)

## Installation & Usage

### Running Locally

```bash
# Quick start with PHP built-in server
php -S localhost:8000

# Then open http://localhost:8000 in your browser
```

### Full Setup (Apache/Nginx)

1. Copy files to your web root (e.g. `/var/www/html/`)
2. If a database is needed, import the `.sql` file into MySQL
3. Copy `config.example.php` to `config.php` and fill in your settings
4. Open the site in your browser

## Origin

This project was originally hosted on SourceForge and has been migrated to GitHub for easier access and collaboration.

- **SourceForge:** [saveofficedata](https://sourceforge.net/projects/saveofficedata/)
- **Migrated with:** [SF2GH Migrator](https://github.com/socrtwo/sf-to-github)

## Contributing

Contributions are welcome! Feel free to:

1. Fork this repository
2. Create a feature branch (`git checkout -b my-feature`)
3. Commit your changes (`git commit -m "Add my feature"`)
4. Push to the branch (`git push origin my-feature`)
5. Open a Pull Request

## License

MIT License — see [LICENSE](LICENSE) for details.