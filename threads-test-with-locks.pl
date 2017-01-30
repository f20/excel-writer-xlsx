#!/usr/bin/env perl

use strict;
use warnings;
use Cwd;
use File::Spec;

use threads;
use Thread::Queue;
my $completion_queue = Thread::Queue->new;

use threads::shared;
my $chdir_lock = 0;
share $chdir_lock;

use Excel::Writer::XLSX;
-d 'test_output' or mkdir 'test_output' or die $!;
my $running_count = 0;
while (1) {
    while ( $running_count > 5 || threads->list(threads::running) > 30 ) {
        my $completion_message = $completion_queue->dequeue;
        --$running_count if $completion_message eq 'Workbooks complete';
        $_->join foreach threads->list(threads::joinable);
    }
    threads->create( \&run_thread );
    ++$running_count;
}

sub make_workbook {
    my ($name) = @_;
    {
        lock $chdir_lock;
        ++$chdir_lock;
    }
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
    {
        lock $chdir_lock;
        --$chdir_lock;
        cond_signal $chdir_lock;
    }
    $workbook;
}

sub run_thread {
    my @books =
      grep { $_; }
      map { make_workbook( 'Thread ' . threads->tid . ' book ' . $_ ); } 1 .. 3;
    $completion_queue->enqueue('Workbooks complete');
    {
        lock $chdir_lock;
        cond_wait $chdir_lock while $chdir_lock > 0;
        $_->DESTROY foreach @books;
        cond_signal $chdir_lock;
    }
    $completion_queue->enqueue('Cleanup complete');
}

