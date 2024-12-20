<?php

declare(strict_types=1);

/*
 * This file is part of package ang3/php-excel-encoder
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace AssoConnect\Serializer\Encoder;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Reader as Readers;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer as Writers;
use PhpOffice\PhpSpreadsheet\Writer\BaseWriter;
use Symfony\Component\Filesystem\Filesystem;
use Symfony\Component\Serializer\Encoder\CsvEncoder;
use Symfony\Component\Serializer\Encoder\DecoderInterface;
use Symfony\Component\Serializer\Encoder\EncoderInterface;
use Symfony\Component\Serializer\Exception\InvalidArgumentException;
use Symfony\Component\Serializer\Exception\NotEncodableValueException;
use Symfony\Component\Serializer\Exception\NotNormalizableValueException;
use Symfony\Component\Serializer\Exception\RuntimeException;

/**
 * @author Joanis ROUANET
 */
class ExcelEncoder implements EncoderInterface, DecoderInterface
{
    /**
     * Formats constants.
     */
    public const XLS = 'xls';
    public const XLSX = 'xlsx';

    /**
     * Context constants.
     */
    public const AS_COLLECTION_KEY = CsvEncoder::AS_COLLECTION_KEY;
    public const FLATTENED_HEADERS_SEPARATOR_KEY = 'flattened_separator_key';
    public const HEADERS_IN_BOLD_KEY = 'headers_in_bold';
    public const HEADERS_HORIZONTAL_ALIGNMENT_KEY = 'headers_horizontal_alignment';
    public const COLUMNS_AUTOSIZE_KEY = 'columns_autosize';
    public const COLUMNS_MAXSIZE_KEY = 'columns_maxsize';


    /**
     * @var string[]
     */
    private static array $formats = [
        self::XLS,
        self::XLSX,
    ];

    /**
     * @var array<string, mixed>
     */
    private array $defaultContext;

    private Filesystem $filesystem;

    /**
     * @param mixed[] $defaultContext
     */
    public function __construct(array $defaultContext = [])
    {
        $baseDefaultContext = [
            self::AS_COLLECTION_KEY => true,
            self::FLATTENED_HEADERS_SEPARATOR_KEY => '.',
            self::HEADERS_IN_BOLD_KEY => true,
            self::HEADERS_HORIZONTAL_ALIGNMENT_KEY => 'center',
            self::COLUMNS_AUTOSIZE_KEY => true,
            self::COLUMNS_MAXSIZE_KEY => 50,
        ];

        $this->defaultContext = array_merge($baseDefaultContext, $defaultContext);
        $this->filesystem = new Filesystem();
    }

    /**
     * {@inheritdoc}.
     */
    public function supportsEncoding(string $format): bool
    {
        return \in_array($format, self::$formats, true);
    }

    /**
     * {@inheritdoc}.
     *
     * @param mixed[] $context
     *
     * @throws InvalidArgumentException   When the format is not supported
     * @throws NotEncodableValueException When data are not valid
     * @throws RuntimeException           When data writing failed
     * @throws PhpSpreadsheetException    On data failure
     */
    public function encode(mixed $data, string $format, array $context = []): string
    {
        if (!is_iterable($data)) {
            throw new NotEncodableValueException(sprintf('Expected data of type iterable, %s given', \gettype($data)));
        }

        $context = $this->normalizeContext($context);
        $spreadsheet = new Spreadsheet();

        $writer = match ($format) {
            self::XLSX => new Writers\Xlsx($spreadsheet),
            self::XLS => new Writers\Xls($spreadsheet),
            default => throw new InvalidArgumentException(sprintf('The format "%s" is not supported', $format)),
        };

        $sheetIndex = 0;
        foreach ($data as $sheetName => $sheetData) {
            $this->processSheetData($sheetName, $sheetData, $spreadsheet, $context, $sheetIndex++);
        }

        return $this->writeToFile($writer, $format);
    }

    /**
     * @param string $sheetName
     * @param iterable<mixed> $sheetData
     * @param Spreadsheet $spreadsheet
     * @param mixed[] $context
     * @param int $sheetIndex
     * @return void
     */
    private function processSheetData(
        string $sheetName,
        iterable $sheetData,
        Spreadsheet $spreadsheet,
        array $context,
        int $sheetIndex
    ): void {
        if ($sheetIndex > 0) {
            $spreadsheet->createSheet($sheetIndex);
        }

        $spreadsheet->setActiveSheetIndex($sheetIndex);
        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->setTitle($sheetName);
        $sheetData = (array) $sheetData;

        foreach ($sheetData as $rowIndex => $cells) {
            if (!is_iterable($cells)) {
                throw new NotEncodableValueException(
                    sprintf(
                        'Expected cells of type "iterable" for data sheet #%d at row #%d, "%s" given',
                        $sheetIndex,
                        $rowIndex,
                        \gettype($cells)
                    )
                );
            }
            $flattened = [];
            $this->flatten($cells, $flattened, $context[self::FLATTENED_HEADERS_SEPARATOR_KEY]);
            $sheetData[$rowIndex] = $flattened;
        }

        $headers = [];
        foreach ($sheetData as $cells) {
            $headers = array_keys($cells);
            break;
        }
        array_unshift($sheetData, $headers);
        $worksheet->fromArray($sheetData, null, 'A1', true);

        $this->applyHeaderStyles($worksheet, $context);
    }

    /**
     * @param mixed[] $context
     */
    private function applyHeaderStyles(Worksheet $worksheet, array $context): void
    {
        $headerLineStyle = $worksheet->getStyle('A1:' . $worksheet->getHighestDataColumn() . '1');
        if ($context[self::HEADERS_HORIZONTAL_ALIGNMENT_KEY]) {
            $alignment = match ($context[self::HEADERS_HORIZONTAL_ALIGNMENT_KEY]) {
                'left' => Alignment::HORIZONTAL_LEFT,
                'center' => Alignment::HORIZONTAL_CENTER,
                'right' => Alignment::HORIZONTAL_RIGHT,
                default => throw new InvalidArgumentException(
                    sprintf(
                        'The value of context key "%s" is not valid (possible values: "left", "center" or "right")',
                        self::HEADERS_HORIZONTAL_ALIGNMENT_KEY
                    )
                ),
            };
            $headerLineStyle->getAlignment()->setHorizontal($alignment);
        }

        if (true === $context[self::HEADERS_IN_BOLD_KEY]) {
            $headerLineStyle->getFont()->setBold(true);
        }

        for ($i = 1; $i <= Coordinate::columnIndexFromString($worksheet->getHighestDataColumn()); ++$i) {
            $worksheet->getColumnDimensionByColumn($i)->setAutoSize($context[self::COLUMNS_AUTOSIZE_KEY]);
        }
        $worksheet->calculateColumnWidths();

        foreach ($worksheet->getColumnDimensions() as $columnDimension) {
            if ($columnDimension->getWidth() > $context[self::COLUMNS_MAXSIZE_KEY]) {
                $columnDimension->setAutoSize(false);
                $columnDimension->setWidth($context[self::COLUMNS_MAXSIZE_KEY]);
            }
        }
    }

    /**
     * @param BaseWriter $writer
     * @return string
     */
    private function writeToFile(BaseWriter $writer, string $format): string
    {
        try {
            $tmpFile = $this->filesystem->tempnam(sys_get_temp_dir(), $format);
            $writer->save($tmpFile);
            $content = (string) file_get_contents($tmpFile);
            $this->filesystem->remove($tmpFile);
        } catch (\Exception $e) {
            throw new RuntimeException(sprintf('Excel encoding failed - %s', $e->getMessage()), 0, $e);
        }
        return $content;
    }

    /**
     * {@inheritdoc}.
     */
    public function supportsDecoding(string $format): bool
    {
        return \in_array($format, self::$formats, true);
    }

    /**
     * @throws NotEncodableValueException When data are not valid
     * @throws InvalidArgumentException   When the format or data not supported
     * @throws RuntimeException           When data reading failed
     * @throws PhpSpreadsheetException    On data failure
     */
    private function loadSpreadsheet(string $tmpFile, string $format): Spreadsheet
    {
        $reader = match ($format) {
            self::XLSX => new Readers\Xlsx(),
            self::XLS => new Readers\Xls(),
            default => throw new InvalidArgumentException(sprintf('The format "%s" is not supported', $format)),
        };

        try {
            return $reader->load($tmpFile);
        } catch (\Exception $e) {
            throw new RuntimeException(sprintf('Excel decoding failed - %s', $e->getMessage()), 0, $e);
        } finally {
            $this->filesystem->remove($tmpFile);
        }
    }

    /**
     * @param mixed[] $data
     * @return mixed[]
     */
    private function transformSheetData(array $data): array
    {
        $labelledRows = [];
        $headers = null;

        foreach ($data as $rowIndex => $cells) {
            $rowIndex = (int) $rowIndex;
            $isHeaderRow = ($headers === null);

            if ($isHeaderRow) {
                $headers = $this->processHeaders($cells);
                continue;
            }

            $labelledRows[$rowIndex - 1] = $this->processCells($headers, $cells);
        }

        return $labelledRows;
    }

    /**
     * Process the header row
     *
     * @param mixed[] $cells
     * @return mixed[]
     */
    private function processHeaders(array $cells): array
    {
        $headers = [];
        foreach ($cells as $key => $value) {
            if (null === $value || '' === $value) {
                continue;
            }
            $headers[$key] = $value;
        }
        return $headers;
    }

    /**
     * Process the data cells
     *
     * @param mixed[] $headers
     * @param mixed[] $cells
     * @return mixed[]
     */
    private function processCells(array $headers, array $cells): array
    {
        $labelledRow = [];
        foreach ($cells as $key => $value) {
            if (\array_key_exists($key, $headers)) {
                $labelledRow[(string) $headers[$key]] = $value;
            } else {
                $labelledRow[''][$key] = $value;
            }
        }
        return $labelledRow;
    }

    /**
     * @param mixed[] $context
     * @return mixed[]
     */
    public function decode(string $data, string $format, array $context = []): array
    {
        $context = $this->normalizeContext($context);
        $tmpFile = (string) tempnam(sys_get_temp_dir(), $format);
        $this->filesystem->dumpFile($tmpFile, $data);

        $spreadsheet = $this->loadSpreadsheet($tmpFile, $format);
        $loadedSheetNames = $spreadsheet->getSheetNames();

        $data = [];
        foreach ($loadedSheetNames as $sheetIndex => $loadedSheetName) {
            $worksheet = $spreadsheet->getSheet($sheetIndex);
            $sheetData = $worksheet->toArray();
            if (0 === \count($sheetData)) {
                continue;
            }

            $data[$loadedSheetName] = !$context[self::AS_COLLECTION_KEY] ?
                $sheetData :
                $this->transformSheetData($sheetData);
        }
        return $data;
    }

    /**
     * @param array<string, mixed> $result
     * @param iterable<mixed, mixed> $data
     *
     * @throws NotNormalizableValueException when a value is not valid
     * @internal
     *
     */
    private function flatten(iterable $data, array &$result, string $keySeparator, string $parentKey = ''): void
    {
        foreach ($data as $key => $value) {
            if (\is_object($value)) {
                $value = get_object_vars($value);
            }

            if (is_iterable($value)) {
                $this->flatten($value, $result, $keySeparator, $parentKey . $key . $keySeparator);

                continue;
            }

            $newKey = $parentKey . $key;

            if (!\is_scalar($value)) {
                throw new NotNormalizableValueException(
                    sprintf('Expected key "%s" of type object, array or scalar, %s given', $newKey, \gettype($value))
                );
            }

            $result[sprintf('="%s"', $newKey)] = match ($value) {
                false => 0,
                true => 1,
                default => $value,
            };
        }
    }

    /**
     * @param mixed[] $context
     * @return mixed[]
     * @internal
     *
     */
    private function normalizeContext(array $context = []): array
    {
        return [
            self::AS_COLLECTION_KEY => (bool) $this->getContextValue($context, self::AS_COLLECTION_KEY),
            self::FLATTENED_HEADERS_SEPARATOR_KEY => (string) $this->getContextValue(
                $context,
                self::FLATTENED_HEADERS_SEPARATOR_KEY
            ),
            self::HEADERS_IN_BOLD_KEY => (bool) $this->getContextValue($context, self::HEADERS_IN_BOLD_KEY),
            self::HEADERS_HORIZONTAL_ALIGNMENT_KEY => (string) $this->getContextValue(
                $context,
                self::HEADERS_HORIZONTAL_ALIGNMENT_KEY
            ),
            self::COLUMNS_AUTOSIZE_KEY => (bool) $this->getContextValue($context, self::COLUMNS_AUTOSIZE_KEY),
            self::COLUMNS_MAXSIZE_KEY => (int) $this->getContextValue($context, self::COLUMNS_MAXSIZE_KEY),
        ];
    }

    /**
     * @param mixed[] $context
     * @internal
     *
     */
    private function getContextValue(array $context, int|string $key): bool|int|float|string|null
    {
        return $context[$key] ?? $this->defaultContext[$key];
    }
}
