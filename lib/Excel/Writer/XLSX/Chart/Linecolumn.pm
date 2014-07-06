package Excel::Writer::XLSX::Chart::Linecolumn;

use 5.008002;
use strict;
use warnings;
use Carp;
use Excel::Writer::XLSX::Chart::Column;

our @ISA     = qw(Excel::Writer::XLSX::Chart::Column);

###############################################################################
#
# new()
#
#
sub new {

    my $class = shift;
    my $self  = Excel::Writer::XLSX::Chart->new(@_);

    $self->{_default_marker} = { type => 'none' };
    $self->{_subtype} ||= 'clustered';
    $self->{_horiz_val_axis} = 0;

    bless $self, $class;
    return $self;
}

##############################################################################
#
# _write_chart_type()
#
# Override the virtual superclass method with a chart specific method.
#
sub _write_chart_type {

    my $self = shift;
    my %args = @_;

    my $series;
    $series = delete $self->{_series};
    $self->{_series} = [ $series->[0] ];
    bless $self, $ISA[0];
    $self->_write_chart_type(%args);

    $self->{_series} = [ @{$series}[ 1 .. $#$series ] ];
    bless $self, __PACKAGE__;

    my @series;
    if ( $args{primary_axes} ) {
        @series = $self->_get_primary_axes_series;
    }
    else {
        @series = $self->_get_secondary_axes_series;
    }

    return unless scalar @series;

    $self->xml_start_tag('c:lineChart');

    # Write the c:grouping element.
    $self->_write_grouping('standard');

    # Write the series elements.
    $self->_write_series($_) for @series;

    # Write the c:dropLines element.
    $self->_write_drop_lines();

    # Write the c:hiLowLines element.
    $self->_write_hi_low_lines();

    # Write the c:upDownBars element.
    $self->_write_up_down_bars();

    # Write the c:marker element.
    $self->_write_marker_value();

    # Write the c:axId elements
    $self->_write_axis_ids(%args);

    $self->xml_end_tag('c:lineChart');

}

1;

