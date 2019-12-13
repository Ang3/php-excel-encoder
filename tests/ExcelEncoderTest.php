<?php

namespace Ang3\Component\Serializer\Encoder\Tests;

use DateTime;
use Ang3\Component\Serializer\Encoder\ExcelEncoder;
use Symfony\Component\Filesystem\Filesystem;
use PHPUnit\Framework\TestCase;

/**
 * @author Joanis ROUANET
 */
class ExcelEncoderTest extends TestCase
{
    /**
     * @var Filesystem
     */
    private $filesystem;

    /**
     * @var ExcelEncoder
     */
    private $encoder;

    /**
     * Set up test.
     */
    public function setUp(): void
    {
        $this->filesystem = new Filesystem();
        $this->encoder = new ExcelEncoder();
    }

    /**
     * @return array
     */
    public function dataEncodingProvider()
    {
        return [
            [ExcelEncoder::XLS],
            [ExcelEncoder::XLSX],
        ];
    }

    /**
     * @dataProvider dataEncodingProvider
     */
    public function testEncode(string $format)
    {
        // Encodage
        $xls = $this->encoder->encode($this->getInitialData(), 'xlsx');

        // Assertions
        $this->assertIsString('string', (string) $xls);
    }

    /**
     * @return array
     */
    public function dataDecodingProvider()
    {
        return [
            [__DIR__.'/Resources/encoded.xls', ExcelEncoder::XLS, $this->getDecodedData('Sheet_0')],
            [__DIR__.'/Resources/encoded.xlsx', ExcelEncoder::XLSX, $this->getDecodedData('Sheet_0')],
            [__DIR__.'/Resources/encoded.csv', ExcelEncoder::SPREADSHEET, $this->getDecodedData()],
            [__DIR__.'/Resources/encoded.unknown', ExcelEncoder::SPREADSHEET, $this->getDecodedData('Sheet_0')],
        ];
    }

    /**
     * @dataProvider dataDecodingProvider
     */
    public function testDecode(string $file, string $format, array $result)
    {
        $this->assertEquals($result, $this->encoder->decode((string) file_get_contents($file), $format));
    }

    /**
     * @internal
     *
     * @return array
     */
    private function getInitialData()
    {
        return [[
            [
                'bool' => false,
                'int' => 1,
                'float' => 1.618,
                'string' => 'Hello',
                'object' => new DateTime('2000-01-01 13:37:00'),
                'array' => [
                    'bool' => true,
                    'int' => 3,
                    'float' => 3.14,
                    'string' => 'World',
                    'object' => new DateTime('2000-01-01 13:37:00'),
                    'array' => [
                        'again',
                    ],
                ],
            ],
        ]];
    }

    /**
     * @internal
     *
     * @return array
     */
    private function getDecodedData(string $sheetName = 'Worksheet')
    {
        return [
            $sheetName => [
                [
                    'bool' => 0,
                    'int' => 1,
                    'float' => 1.618,
                    'string' => 'Hello',
                    'object.date' => '2000-01-01 13:37:00.000000',
                    'object.timezone_type' => 3,
                    'object.timezone' => 'Europe/Berlin',
                    'array.bool' => 1,
                    'array.int' => 3,
                    'array.float' => 3.14,
                    'array.string' => 'World',
                    'array.object.date' => '2000-01-01 13:37:00.000000',
                    'array.object.timezone_type' => 3,
                    'array.object.timezone' => 'Europe/Berlin',
                    'array.array.0' => 'again',
                ],
            ],
        ];
    }
}
