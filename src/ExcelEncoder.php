<?php

namespace Ang3\Component\Serializer\Encoder;

use Exception;
use Symfony\Component\Serializer\Encoder\DecoderInterface;
use Symfony\Component\Serializer\Encoder\EncoderInterface;
use Symfony\Component\Serializer\Exception\InvalidArgumentException;
use Symfony\Component\Serializer\Exception\NotEncodableValueException;
use Symfony\Component\Serializer\Exception\NotNormalizableValueException;
use Symfony\Component\Serializer\Exception\RuntimeException;
use Symfony\Component\Filesystem\Filesystem;
use PhpOffice\PhpSpreadsheet\Writer as Writers;
use PhpOffice\PhpSpreadsheet\Reader as Readers;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

/**
 * @author Joanis ROUANET
 */
class ExcelEncoder implements EncoderInterface, DecoderInterface
{
    /**
     * Formats constants.
     */
    const XLS = 'xls';
    const XLSX = 'xlsx';
    const SPREADSHEET = 'spreadsheet';

    /**
     * Context constants.
     */
    const NB_HEADERS_ROW_KEY = 'nb_headers_row';
    const FLATTENED_HEADERS_SEPARATOR_KEY = 'flattened_separator_key';
    const HEADERS_IN_BOLD_KEY = 'headers_in_bold';
    const HEADERS_HORIZONTAL_ALIGNMENT_KEY = 'headers_horizontal_alignment';
    const COLUMNS_AUTOSIZE_KEY = 'columns_autosize';
    const COLUMNS_MAXSIZE_KEY = 'columns_maxsize';
    const CSV_DELIMITER_KEY = 'csv_delimiter';
    const CSV_ENCLOSURE_KEY = 'csv_enclosure';

    /**
     * @static
     *
     * @var array
     */
    private static $formats = [
        self::XLS,
        self::XLSX,
    ];

    /**
     * @var array
     */
    private $defaultContext = [
        self::NB_HEADERS_ROW_KEY => 1,
        self::FLATTENED_HEADERS_SEPARATOR_KEY => '.',
        self::HEADERS_IN_BOLD_KEY => true,
        self::HEADERS_HORIZONTAL_ALIGNMENT_KEY => 'center',
        self::COLUMNS_AUTOSIZE_KEY => true,
        self::COLUMNS_MAXSIZE_KEY => 50,
        self::CSV_DELIMITER_KEY => ';',
        self::CSV_ENCLOSURE_KEY => '"',
    ];

    /**
     * @var Filesystem
     */
    private $filesystem;

    public function __construct(array $defaultContext = [])
    {
        $this->defaultContext = array_merge($this->defaultContext, $defaultContext);
        $this->filesystem = new Filesystem();
    }

    /**
     * {@inheritdoc}.
     *
     * @throws InvalidArgumentException   When the format is not supported
     * @throws NotEncodableValueException When data are not valid
     * @throws RuntimeException           When data writing failed
     */
    public function encode($data, $format, array $context = [])
    {
        // Si les données ne sont pas sous forme itérable
        if (!is_iterable($data)) {
            throw new NotEncodableValueException(sprintf('Expected data of type iterable, %s given', gettype($data)));
        }

        // Normalisation du contexte
        $context = $this->normalizeContext($context);

        // Création du gestionnaire de classeur
        $spreadsheet = new Spreadsheet();

        // Création du writer selon le format
        switch ($format) {
            // Excel 2007
            case self::XLSX:
                $writer = new Writers\Xlsx($spreadsheet);
            break;

            // Excel 2003
            case self::XLS:
                $writer = new Writers\Xls($spreadsheet);
            break;

            default:
                throw new InvalidArgumentException(sprintf('The format "%s" is not supported', $format));
            break;
        }

        // Initialisation de l'index de la feuille
        $sheetIndex = 0;

        // Pour chaque onglet dans les données
        foreach ($data as $sheetName => $sheetData) {
            // Si les données de l'onglet ne sont pas sous forme itérable
            if (!is_iterable($sheetData)) {
                throw new NotEncodableValueException(sprintf('Expected data of sheet #%d of type "iterable", "%s" given', $sheetName, gettype($sheetData)));
            }

            // Si le nom de la feuille est un entier
            if ($sheetIndex === $sheetName) {
                // Enregistrement du titre
                $sheetName = sprintf('Sheet_%d', $sheetIndex);
            }

            // Assignation de l'onglet
            $spreadsheet->setActiveSheetIndex($sheetIndex);

            // Récupération de l'onglet courant
            $worksheet = $spreadsheet->getActiveSheet();

            // Enregistrement du titre
            $worksheet->setTitle($sheetName);

            // Typage des données en tableau
            $sheetData = (array) $sheetData;

            // Pour chaque ligne
            foreach ($sheetData as $rowIndex => $cells) {
                // Si les données ne sont pas sous forme itérable
                if (!is_iterable($cells)) {
                    throw new NotEncodableValueException(sprintf('Expected cells of type "iterable" for data sheet #%d at row #%d, "%s" given', $sheetIndex, $rowIndex, gettype($cells)));
                }

                // Initialisation du résultat de l'applatissement
                $flattened = [];

                // On aplatit avec le tableau initialisé
                $this->flatten($cells, $flattened, $context[self::FLATTENED_HEADERS_SEPARATOR_KEY]);

                // Mise-à-jour de la valeur par l'applatissement
                $sheetData[$rowIndex] = $flattened;
            }

            // Initialisation des entêtes
            $headers = [];

            // Si on a pas encore d'entêtes
            if (!$headers) {
                foreach ($sheetData as $rowIndex => $cells) {
                    // Récupération des entêtes par la clé des celulles
                    $headers = array_keys($cells);

                    // On stoppe
                    break;
                }
            }

            // On ajoute la entêtes en début de données
            array_unshift($sheetData, $headers);

            // Importation des données depuis le tablau de données
            $worksheet->fromArray((array) $sheetData, null, 'A1', true);

            // Récupération des styles des entêtes
            $headerLineStyle = $worksheet->getStyle('A1:'.$worksheet->getHighestDataColumn().'1');

            // Si on a une valeur pour l'alignement horizontal des entêtes
            if ($context[self::HEADERS_HORIZONTAL_ALIGNMENT_KEY]) {
                // Selon l'alignement souhaité
                switch ($context[self::HEADERS_HORIZONTAL_ALIGNMENT_KEY]) {
                    case 'left':
                        $alignment = Alignment::HORIZONTAL_LEFT;
                    break;
                    case 'center':
                        $alignment = Alignment::HORIZONTAL_CENTER;
                    break;
                    case 'right':
                        $alignment = Alignment::HORIZONTAL_RIGHT;
                    break;
                    default:
                        throw new InvalidArgumentException(sprintf('The value of context key "%s" is not valid (possible values: "left", "center" or "right")', self::HEADERS_HORIZONTAL_ALIGNMENT_KEY));
                    break;
                }

                // On centre les entêtes
                $headerLineStyle
                    ->getAlignment()
                    ->setHorizontal($alignment)
                ;
            }

            // Si on souhaite mettre les entêtes en gras
            if (true === $context[self::HEADERS_IN_BOLD_KEY]) {
                // Mise en gras des entêtes
                $headerLineStyle
                    ->getFont()
                    ->setBold(true)
                ;
            }

            // Pour chaque colonne contenant des données
            for ($i = 1; $i <= Coordinate::columnIndexFromString($worksheet->getHighestDataColumn()); ++$i) {
                // On dimensionne automatiquement la colonne
                $worksheet
                    ->getColumnDimensionByColumn($i)
                    ->setAutoSize($context[self::COLUMNS_AUTOSIZE_KEY])
                ;
            }

            // On lance le calcul des dimensions des colonnes
            $worksheet->calculateColumnWidths();

            // Pour chaque dimension de colonne définit
            foreach ($worksheet->getColumnDimensions() as $columnDimension) {
                // Récupération de la taille de la colonne
                $colWidth = $columnDimension->getWidth();

                // Si la taille de la colonne dépasse la taille maximale des colonnes
                if ($colWidth > $context[self::COLUMNS_MAXSIZE_KEY]) {
                    // Désactivation de l'option d'autosize
                    $columnDimension->setAutoSize(false);

                    // Enregistrement de la taille maximale pour la colonne
                    $columnDimension->setWidth($context[self::COLUMNS_MAXSIZE_KEY]);
                }
            }
        }

        try {
            // Création du fichier temporaire de destination
            $tmpFile = $this->filesystem->tempnam(sys_get_temp_dir(), $format);

            // On écrit classeur en XLS dans le fichier temporaire
            $writer->save($tmpFile);

            // Retour de la lecture du fichier temporaire
            $content = file_get_contents($tmpFile);

            // Suppression du fichier temporaire
            $this->filesystem->remove($tmpFile);
        } catch (Exception $e) {
            throw new RuntimeException(sprintf('Excel encoding failed - %s', $e->getMessage()), 0, $e);
        }

        // Retour du contenu Excel
        return $content;
    }

    /**
     * {@inheritdoc}.
     */
    public function supportsEncoding($format)
    {
        return in_array($format, self::$formats);
    }

    /**
     * {@inheritdoc}.
     *
     * @throws NotEncodableValueException When data are not valid
     * @throws InvalidArgumentException   When the format or data not supported
     * @throws RuntimeException           When data reading failed
     */
    public function decode($data, $format, array $context = [])
    {
        // Si les données ne sont pas scalaire
        if (!is_scalar($data)) {
            throw new NotEncodableValueException(sprintf('Expected data of type scalar, %s given', gettype($data)));
        }

        // Normalisation du contexte
        $context = $this->normalizeContext($context);

        // Création du fichier temporaire de destination
        $tmpFile = tempnam(sys_get_temp_dir(), $format);

        // On insère les données dans le fichier temporaire
        $this->filesystem->dumpFile($tmpFile, $data);

        // Création du lecteur selon le format
        switch ($format) {
            // Excel 2007
            case self::XLSX:
                $reader = new Readers\Xlsx();
            break;

            // Excel 2003
            case self::XLS:
                $reader = new Readers\Xls();
            break;

            // Pas de format spécifique
            case self::SPREADSHEET:
                try {
                    // Création d'un lecteur pour le fichier temporaire
                    $reader = IOFactory::createReaderForFile($tmpFile);
                } catch (Exception $e) {
                    throw new NotEncodableValueException(sprintf('Unable to determine data format - %s', $e->getMessage()), 0, $e);
                }
            break;

            default:
                throw new InvalidArgumentException(sprintf('The format "%s" is not supported', $format));
            break;
        }

        // Si le lecteur est pour du CSV
        if ($reader instanceof Readers\Csv) {
            // On configure le délimiter et l'enclosure
            $reader
                ->setDelimiter($context[self::CSV_DELIMITER_KEY])
                ->setEnclosure($context[self::CSV_ENCLOSURE_KEY])
            ;
        }

        try {
            // Chargement du classeur
            $spreadsheet = $reader->load($tmpFile);

            // Suppression du fichier temporaire
            $this->filesystem->remove($tmpFile);
        } catch (Exception $e) {
            throw new RuntimeException(sprintf('Excel decoding failed - %s', $e->getMessage()), 0, $e);
        }

        // Récupération du nom des onglets
        $loadedSheetNames = $spreadsheet->getSheetNames();

        // Initialisation des données
        $data = [];

        // Pour chaque nom d'onglet
        foreach ($loadedSheetNames as $sheetIndex => $loadedSheetName) {
            // Récupération de la feuille
            $worksheet = $spreadsheet->getSheet($sheetIndex);

            // Récupération des données
            $sheetData = $worksheet->toArray();

            // Si pas de données
            if (0 === count($sheetData)) {
                // Onglet suivant
                continue;
            }

            // Si on a pas d'entêtes
            if (0 === $context[self::NB_HEADERS_ROW_KEY]) {
                // Enregistrement des lignes de la feuilles dans les données
                $data[$loadedSheetName] = $sheetData;

                // Feuille suivante
                continue;
            }

            // Initialisation des lignes avec leur entête
            $labelledRows = [];

            // Initialisation des entêtes
            $headers = [];

            // Initialisation du nombre de ligne d'entêtes enresgitrées
            $headerRowsCount = 0;

            // Pour chaque ligne de données
            foreach ($sheetData as $rowIndex => $cells) {
                // Typage de l'index de la ligne en entier
                $rowIndex = (int) $rowIndex;

                // Si le nombre de lignes d'entêtes attendus est supérieur au nombre de lignes d'entête enregistrés
                if ($context[self::NB_HEADERS_ROW_KEY] > $headerRowsCount) {
                    // Pour chaque valeur de la ligne
                    foreach ($cells as $key => $value) {
                        // Si pas de valeur
                        if (null === $value || '' === $value) {
                            // Entête suivant
                            continue;
                        }

                        // Enregistrement de la valeur de l'entête sur la colonne
                        $headers[$key] = $value;

                        // Suppression de la valeur
                        unset($sheetData[$rowIndex][$key]);
                    }

                    // Incrémentation du nombre de lignes d'entêtes enregistrées
                    ++$headerRowsCount;

                    // Ligne suivante
                    continue;
                }

                // Pour chaque valeur de la ligne
                foreach ($cells as $key => $value) {
                    // Si on a un entête pour cette valeur
                    if (array_key_exists($key, $headers)) {
                        // Enregistrement de la valeur
                        $labelledRows[$rowIndex - 1][(string) $headers[$key]] = $value;
                    } else {
                        // Enregistrement de la valeur sans entête selon sa clé
                        $labelledRows[$rowIndex - 1][''][$key] = $value;
                    }

                    // Suppression de la valeur
                    unset($sheetData[$rowIndex][$key]);
                }

                // Suppression de la ligne dans les données originelles
                unset($sheetData[$rowIndex]);
            }

            // Enregistrement des lignes dans les données
            $data[$loadedSheetName] = $labelledRows;

            // Incrémentation de l'index de feuille
            ++$sheetIndex;
        }

        // Retour des données
        return $data;
    }

    /**
     * {@inheritdoc}.
     */
    public function supportsDecoding($format)
    {
        return in_array($format, array_merge(self::$formats, [self::SPREADSHEET]));
    }

    /**
     * @throws NotNormalizableValueException when a value is not valid
     *
     * @return array
     */
    public function flatten(iterable $data, array &$result = [], string $keySeparator, string $parentKey = '')
    {
        // Pour chaque valeur
        foreach ($data as $key => $value) {
            // Si la valeur est un objet
            if (is_object($value)) {
                // Normalisation de l'objet
                $value = get_object_vars($value);
            }

            // Si on a encore une valeur itérable
            if (is_iterable($value)) {
                // On aplatit le tableau par récursivité
                $this->flatten($value, $result, $keySeparator, $parentKey.$key.$keySeparator);

                // Valeur suivante
                continue;
            }

            // Définition de la sous-clé
            $newKey = $parentKey.$key;

            // Si la valeur n'est toujours pas scalaire
            if (!is_scalar($value)) {
                throw new NotNormalizableValueException(sprintf('Expected key "%s" of type object, array or scalar, %s given', $newKey, gettype($value)));
            }

            // Enregistrement de la clé en faisant attention aux clé scalaires et aux valeurs booléennes
            $result[sprintf('="%s"', $newKey)] = false === $value ? 0 : (true === $value ? 1 : $value);
        }
    }

    /**
     * @internal
     *
     * @return array
     */
    private function normalizeContext(array $context = [])
    {
        return [
            self::NB_HEADERS_ROW_KEY => (int) $this->getContextValue($context, self::NB_HEADERS_ROW_KEY),
            self::FLATTENED_HEADERS_SEPARATOR_KEY => (string) $this->getContextValue($context, self::FLATTENED_HEADERS_SEPARATOR_KEY),
            self::HEADERS_IN_BOLD_KEY => (bool) $this->getContextValue($context, self::HEADERS_IN_BOLD_KEY),
            self::HEADERS_HORIZONTAL_ALIGNMENT_KEY => (string) $this->getContextValue($context, self::HEADERS_HORIZONTAL_ALIGNMENT_KEY),
            self::COLUMNS_AUTOSIZE_KEY => (bool) $this->getContextValue($context, self::COLUMNS_AUTOSIZE_KEY),
            self::COLUMNS_MAXSIZE_KEY => (int) $this->getContextValue($context, self::COLUMNS_MAXSIZE_KEY),
            self::CSV_DELIMITER_KEY => (string) $this->getContextValue($context, self::CSV_DELIMITER_KEY),
            self::CSV_ENCLOSURE_KEY => (string) $this->getContextValue($context, self::CSV_ENCLOSURE_KEY),
        ];
    }

    /**
     * @internal
     *
     * @return mixed
     */
    private function getContextValue(array $context = [], $key)
    {
        return $context[$key] ?? $this->defaultContext[$key];
    }
}
