<?php

namespace Ang3\Component\Serializer\Encoder;

use DateTIme;
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
	 * @var array
	 */
	private $data = [];

	/**
	 * Set up test.
	 */
	public function setUp() : void
	{
		$this->filesystem = new Filesystem;
		$this->encoder = new ExcelEncoder;
		$this->data = [
			'Sheet_0' => [
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
					],
				]
			]
		];
	}

	public function testDecodeXlsx()
	{
		$this->assertEquals([
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
				]
			]
		], $this->encoder->decode(file_get_contents(__DIR__.'/Resources/encoded.xlsx'), 'xlsx'));
	}

	/**
	 * @depends testDecodeXlsx
	 */
	public function testEncodeXslx()
	{
		$this->assertTrue(file_get_contents(__DIR__.'/Resources/encoded.xlsx') === $this->encoder->encode($this->data, 'xlsx'));
	}
}