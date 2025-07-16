<?php
use Illuminate\Support\Arr;
use Illuminate\Support\Facades\Route;

Route::middleware(['web'])->prefix('document')->name('document.')->group(function() {
    foreach (Arr::get(config('documents'), 'documents', []) as $doc) {
        $name = strtolower((new ReflectionClass($doc))->getShortName());
        Route::any("/{$name}/{id}", $doc)->name($name);
    }
});