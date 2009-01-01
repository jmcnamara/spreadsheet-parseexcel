package Test::Example;
use strict;
use warnings;

use base 'Exporter';
our @EXPORT_OK = ('test_example_do');
use Test::Builder;
use Carp;


my $T = Test::Builder->new;

sub test_example_do {
    my %args = @_;

    #chdir $args{dir} if $args{dir}; # TODO: chdir back at the end?

    close STDERR;
    close STDOUT;
    
    unshift @INC, '../blib/lib'; # that should of course depend on the $args{dir}

    my $err = '';
    open STDERR, '>', \$err or die;
    my $out = '';
    open STDOUT, '>', \$out or die;
    if ($args{argv}) {
        @ARGV = @{ $args{argv} };
        $T->diag("@ARGV");
    }
    #$T->diag("Command: $args{dir}/$args{script}");
    #$T->diag(`pwd`);
    if (not defined do("$args{dir}/$args{script}")) {
        $T->diag("Could not run example: $!");
    }
    # TODO: eval ? check $@ ?
    #system "$^X $args{dir}/$args{script} @ARGV"; 
    @ARGV = ();
    close STDERR;
    close STDOUT;
    #diag $out;
    #my $err = slurp('err');
    $T->is_eq($err, '', "stderr is empty when running $args{script}");
    if (defined $args{stdout}) {
        my @expected_out = slurp("$args{dir}/$args{stdout}");
        #my @out = slurp('out');
        my @out = $out =~ /^.*\n/mg;
        #$T->is_num(scalar(@out), scalar(@expected_out), "STDOUT rownumber is correct");
        compare_lists(\@out, \@expected_out);
        #$T->is_deeply(\@out, \@expected_out, "stdout when running $t->{in}");
    }
    else {
        $T->is_eq($out, "", "STDOUT is empty");
    }
}

# TODO: replace by List::Compare?
sub compare_lists {
    my ($result, $expected) = @_;
    croak 'need two params' if not ($result and $expected);
    croak 'not array refs' if not (ref($result) eq 'ARRAY' and ref($expected) eq 'ARRAY');
    if (@$expected != @$result) {
        $T->ok(0);
        $T->diag("Lists are not the same length. Expected is " 
            . @$expected . " long, while received is " . @$result . " long");
        return;
    }
    foreach my $i (0..@$result-1) {
        if ($result->[$i] ne $expected->[$i]) {
            $T->ok(0);
            $T->diag("In row $i");
            $T->diag("Expected: $expected->[$i]");
            $T->diag("Received: $result->[$i]");
            return;
        }
    }
    $T->ok(1);
}


sub slurp {
    my ($file) = @_;
    open my $fh, '<', $file or die "Could not open $file: $!";
    if (wantarray) {
        return <$fh>;
    } else {
        local $/ = undef;
        return <$fh>;
    }
}


1;

__END__
=head1 NAME

Test::Example - Check if all the examples in the distribution work correctly

=head1 SYNOPSIS

    use Test::Example;
    test_all_examples();

or

    use Test::Example;
    foreach my $file (glob 'myexamples/*.plx') {
        test_example(
            dir    => 'myexamples',
            script => $file,
            stdout => "stdout/$file",
            stderr => "stderr/$file",
        );
    }


=head1 METHODS


=head2 test_all_examples

Goes over all the .pl files in the eg/ examples/ /sample/  (...?) 
directories runs each one of the scripts using L<test_example>.
Options given to test_example are:

    test_example(
        dir    => 'eg',                  # the name of the relevant directory
        script => 'scriptname.pl',       # the name of the current .pl file
        stdin  => 'scriptname.pl_stdin',
        stdout => 'scriptname.pl_stdout',
        stderr => 'scriptname.pl_stderr',
    );


=head2 test_all_examples_do

The same as test_all_examples but 

=head2  test_example

    test_example(
        dir     => 'myexamples',
        script  => 'doit.pl',
        stdin   => 'file_providing_stdin',
        stdout  => 'file_listing_expected_output_of_doit',
        stderr  => 'file_listing_expected_errors_of_doit',
        argv    => ['command', 'line', 'arguments'],
    );

Before running doit.pl chdirs into the 'myexamples' directory.
doit.pl is executed using system. The list of values provided
as argv are supplied as command line parameters.
Its STDIN is redirected from the file that is given as 'stdin'.
Its STDOUT and STDERR are captured.

In short, something like this:

    chdir 'myexamples';
    system("$h{script} @{ $h{argv} } < $h{stdin} > temp_out 2> temp_err"); 

Once the script finished the content of temp_out is compared to
the expeced output and the content of temp_err to the expected errors.

If no 'stderr' key provided then the expectation is that nothing
will be printed to STDERR.

=head2 test_example_do

The same as L</test_example> but instead of using C<system> to run the external
script it will use C<do 'scriptname.pl'>

=cut

