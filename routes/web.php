<?php

use App\Http\Controllers\ReaderController;
use Illuminate\Support\Facades\Route;

Route::get('/excel', [ReaderController::class, 'store'])->name('reader.store');

Route::inertia('/', 'Welcome')->name('home');

Route::middleware(['auth', 'verified'])->group(function () {
    Route::inertia('dashboard', 'Dashboard')->name('dashboard');
});

require __DIR__.'/settings.php';
