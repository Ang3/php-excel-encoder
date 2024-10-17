<?php

declare(strict_types=1);

/*
 * This file is part of package assoconnect/php-excel-encoder
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace AssoConnect\Serializer\Encoder\Tests;

use AssoConnect\Serializer\Encoder\ExcelEncoder;
use PHPUnit\Framework\TestCase;

/**
 * @author Joanis ROUANET
 *
 * @internal
 *
 * @covers \AssoConnect\Serializer\Encoder\ExcelEncoder
 */
final class ExcelEncoderTest extends TestCase
{
    private ExcelEncoder $encoder;

    /**
     * Set up test.
     */
    protected function setUp(): void
    {
        $this->encoder = new ExcelEncoder();
    }

    public static function dataEncodingProvider(): array
    {
        return [
            [ExcelEncoder::XLS],
            [ExcelEncoder::XLSX],
        ];
    }

    /**
     * @dataProvider dataEncodingProvider
     */
    public function testEncode(string $format): void
    {
        $contents = $this->encoder->encode($this->getInitialData(), $format);
        static::assertIsString('string', $contents);
    }

    public static function dataDecodingProvider(): array
    {
        return [
            [
                __DIR__ . '/Resources/encoded.xls',
                ExcelEncoder::XLS,
                [
                    'Sheet_0' => [
                        [
                            'bool' => 0.0,
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
                    'Feuil1' => [
                        [
                            'bool' => 0.0,
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
                ],
            ],
            [
                __DIR__ . '/Resources/encoded.xlsx',
                ExcelEncoder::XLSX,
                [
                    'Sheet_0' => [
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
                    'Feuil1' => [
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
                ],
            ],
        ];
    }

    /**
     * @dataProvider dataDecodingProvider
     */
    public function testDecode(string $file, string $format, array $result): void
    {
        static::assertSame($result, $this->encoder->decode((string)file_get_contents($file), $format));
    }

    /**
     * @internal
     */
    private function getInitialData(): array
    {
        return [
            'Sheet_0' => [
                [
                    'bool' => false,
                    'int' => 1,
                    'float' => 1.618,
                    'string' => 'Hello',
                    'object' => new \DateTime('2000-01-01 13:37:00'),
                    'array' => [
                        'bool' => true,
                        'int' => 3,
                        'float' => 3.14,
                        'string' => 'World',
                        'object' => new \DateTime('2000-01-01 13:37:00'),
                        'array' => [
                            'again',
                        ],
                    ],
                ],
            ],
            'Feuil1' => [
                [
                    'bool' => false,
                    'int' => 1,
                    'float' => 1.618,
                    'string' => 'Hello',
                    'object' => new \DateTime('2000-01-01 13:37:00'),
                    'array' => [
                        'bool' => true,
                        'int' => 3,
                        'float' => 3.14,
                        'string' => 'World',
                        'object' => new \DateTime('2000-01-01 13:37:00'),
                        'array' => [
                            'again',
                        ],
                    ],
                ],
            ],
        ];
    }
}
