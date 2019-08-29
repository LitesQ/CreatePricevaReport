<?php

	set_time_limit(0);
	ini_set('max_execution_time', 0);

	//Parser::$base_filename = getcwd().'..\inout\base.csv';
	Parser::$base_filename = '..\in\base.csv';
	Parser::$additionals_filename = getcwd().'\additionals.csv';
	Parser::start();

	class Parser {
		public static $base_filename = 'base.csv';
		public static $additionals_filename = 'additionals.csv';
		public static function start() {
			$urls = self::get_urls();
			$base_values = array();
			$additionals_values = array();
			foreach ($urls as $url) {
				echo $url."\n";
				$html = file_get_contents($url);
				$available = self::check_html($html);
				if ($available) {
					array_push($base_values, self::parse_base($html));
					array_push($additionals_values, self::parse_additionals($html));
				}
			}
			$base_keys = array('id', 'name', 'price');
			$additionals_keys = array('id', 'name', 'price');
			for ($i = 1; $i < 11; $i++) {
				$base_keys = array_merge($base_keys, array('seller #'.$i, 'price #'.$i, 'bonus #'.$i, 'bonus, % #'.$i));
				$additionals_keys = array_merge($additionals_keys, array('seller #'.$i, 'price #'.$i));
			}
			array_unshift($base_values, $base_keys);
			array_unshift($additionals_values, $additionals_keys);
			self::save_to_csv(self::$base_filename, $base_values, false, 'windows-1251');
			self::save_to_csv(self::$additionals_filename, $additionals_values, false, 'windows-1251');
		}
		private static function get_urls() {
			$contents = file_get_contents(getcwd().'\urls.txt');
			$urls = explode("\n", $contents);
			return $urls;
		}
		private static function check_html($html) {
			return strpos($html, '_GOODS_.googleAnalyticsProductData') > -1;
		}
		private static function parse_base($html) {
			$item = array(
				'id'			=> self::text_between($html, '_GOODS_.googleAnalyticsProductData[\'', '\''),
				'name'			=> html_entity_decode(self::text_between($html, '<h1 itemprop="name">', '</h1>')),
				'price'			=> self::to_value(self::text_between($html, 'dimension34: \'', '\'')),
			);
			$blocks = self::get_array($html, ':is="buyButtonComponent"', '</component>');
			for ($i = 0; $i < sizeof($blocks); $i++) {
				$block = $blocks[$i];
				$item['seller ('.($i + 1).')'] = html_entity_decode(self::text_between($block, 'shop-name="', '"'));
				$item['price ('.($i + 1).')'] = self::to_value(self::text_between($block, 'product-price="', '"'));
				$item['bonus, value ('.($i + 1).')'] = self::to_value(self::text_between($block, 'bonus-value="', '"'));
				$item['bonus, % ('.($i + 1).')'] = self::to_value(self::text_between($block, 'bonus-percent="', '"') / 100);
			}
			return $item;
		}
		private static function parse_additionals($html) {
			$item = array(
				'id'			=> self::text_between($html, '_GOODS_.googleAnalyticsProductData[\'', '\''),
				'name'			=> html_entity_decode(self::text_between($html, '<h1 itemprop="name">', '</h1>')),
				'price'			=> self::to_value(self::text_between($html, 'dimension34: \'', '\'')),
			);
			$json = self::text_between($html, ':bpg2-data="', '"');
			if ($json) {
				$json = html_entity_decode($json);
				$competitors = json_decode($json, true);
				for ($i = 0; $i < sizeof($competitors); $i++) {
					$competitor = $competitors[$i];
					$item['seller ('.($i + 1).')'] = $competitor['company'];
					$item['price ('.($i + 1).')'] = self::to_value($competitor['price']);
				}
			}
			return $item;
		}
		private static function text_between($source_text, $start_text = '', $end_text = '', $counting = 1) {
			$counting = ($counting == 0) ? 1 : $counting;
			$start_pos = -1;
			if ($start_text == '') {
				$start_pos = 0;
			} else {
				for ($i = 1; $i <= $counting; $i++) {
					$start_pos = strpos($source_text, $start_text, $start_pos + 1);
					if ($start_pos === false) break;
				}
			}
			if ($start_pos === false) {
				return '';
			} else {
				$start_pos = $start_pos + strlen($start_text);
			}
			$end_pos = false;
			if ($end_text == '') {
				$result = substr($source_text, $start_pos);
			} else {
				$end_pos = strpos($source_text, $end_text, $start_pos);
				if ($end_pos === false) {
					$result = substr($source_text, $start_pos);
				} else {
					$result = substr($source_text, $start_pos, $end_pos - $start_pos);
				}
			}
			return $result;
		}
		private static function get_array($source_text, $start_text, $end_text) {
			$result = array();
			$start_pos = 0;
			$end_pos = 0;
			for ($i = 0; $i <= strlen($source_text); $i++) {
				$start_pos = strpos($source_text, $start_text, $i);
				if ($start_pos === false) {
					break;
				} else {
					$start_pos = $start_pos + strlen($start_text);
				}
				$end_pos = strpos($source_text, $end_text, $start_pos);
				if ($end_pos === false) {
					break;
				} else {
					array_push($result, substr($source_text, $start_pos, $end_pos - $start_pos));
					$i = $end_pos;
				}
			}
			return $result;
		}
		private static function save_to_csv($filename, $values, $headings = false, $charset = 'utf-8') {
			if ($headings) {
				$headers = array();
				foreach ($values[0] as $key => $value) {
					$headers[$key] = $key;
				}
				array_unshift($values, $headers);
			}
			if ($charset !== 'utf-8') {
				array_walk_recursive($values, function(&$item) use ($charset) {
					$item = mb_convert_encoding($item, $charset, 'utf-8');
				}, $charset);
			}
			$file = fopen($filename, 'w');
			foreach ($values as $value) {
				fputcsv($file, $value, ';');
			}
			fclose($file);
		}
		private static function to_value($value) {
			$value = str_replace('.', ',', $value);
			return $value;
		}
	}
?>