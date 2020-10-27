#!/usr/bin/env perl

use strict;
use warnings;
use Cwd;
use File::Spec;

use threads;
use Thread::Queue;
my $completion_queue = Thread::Queue->new;

use Excel::Writer::XLSX;
-d 'test_output' or mkdir 'test_output' or die $!;
while (1) {
    while ( threads->list(threads::running) > 5 ) {  # maximum number of workers
        $completion_queue->dequeue;  # wait for one of them to report completion
        $_->join foreach threads->list(threads::joinable);
    }
    threads->create( \&run_thread );    # start a new worker
}

sub make_workbook {
    my ($name) = @_;
    my $workbook =
      Excel::Writer::XLSX->new(
        File::Spec->catfile( 'test_output', $name . '.xlsx' ) );
    if ($workbook) {
        my $worksheet1 = $workbook->add_worksheet();
        $workbook->close;
    }
    else {
        warn "Failed to create workbook; cwd = " . getcwd() . "\n";
    }
    $workbook;
}

sub run_thread {
    map { make_workbook( 'Thread ' . threads->tid . ' book ' . $_ ); } 1 .. 3;
    $completion_queue->enqueue(1);
}

