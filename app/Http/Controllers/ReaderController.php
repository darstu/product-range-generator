<?php

namespace App\Http\Controllers;

use App\Services\SpreadsheetTransformService;

class ReaderController extends Controller
{
    public function store(SpreadsheetTransformService $service)
    {
        $service->transform(
            sourcePath: "files/DB2.xlsx",
            outputPath: storage_path('app/MyExcel.xlsx')
        );

        return response()->json(['status' => 'ok']);
    }
}