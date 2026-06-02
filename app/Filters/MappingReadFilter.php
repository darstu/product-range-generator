<?php

namespace App\Filters;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

class MappingReadFilter implements IReadFilter
{
    public function __construct(
        private array $columns
    ) {}

    public function readCell(
        $columnAddress,
        $row,
        $worksheetName = ''
    ): bool {
        return in_array($columnAddress, $this->columns, true);
    }
}