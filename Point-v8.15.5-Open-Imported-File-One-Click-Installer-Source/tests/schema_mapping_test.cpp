#include "point_core.h"

#include <cassert>
#include <filesystem>
#include <fstream>
#include <iostream>

int main() {
    const auto root = std::filesystem::temp_directory_path() /
        "point-schema-mapping-test";
    std::filesystem::create_directories(root);
    const auto users = root / "Users.csv";
    const auto devices = root / "Devices.csv";
    {
        std::ofstream out(users);
        out << "Employee ID,Username,Email,Display Name,Manager Employee ID\n"
               "E1001,speela,sunil@example.com,Peela Sunil Kumar,\n"
               "E1002,jsmith,john@example.com,Smith John,E1001\n";
    }
    {
        std::ofstream out(devices);
        out << "Associate Number,Computer Name\n"
               "E1001,PC-001\nE1002,PC-002\n";
    }

    point::Engine unmapped;
    unmapped.load_files({users, devices});
    assert(unmapped.relationships().empty());

    point::Engine mapped;
    mapped.set_field_synonyms({
        {"Employee ID", {"EID", "Associate Number", "Worker ID"}}
    });
    mapped.load_files({users, devices});
    assert(mapped.relationships().size() == 1);
    point::QueryRequest request;
    request.lookup_field = "Employee ID";
    request.lookup_value = "E1001";
    request.output_fields = {"Display Name", "Computer Name"};
    const auto result = mapped.query(request);
    assert(result.rows.size() == 1);
    assert(result.rows[0][0] == "Peela Sunil Kumar");
    assert(result.rows[0][1] == "PC-001");

    point::Engine incremental;
    incremental.set_field_synonyms({
        {"Employee ID", {"EID", "Associate Number", "Worker ID"}}
    });
    std::size_t reused_files = 0;
    incremental.load_files_incremental(
        {users, devices}, &mapped,
        [&](std::size_t, std::size_t,
            const std::filesystem::path&, bool reused) {
            if (reused) ++reused_files;
        });
    assert(reused_files == 2);
    assert(incremental.relationships().size() == 1);
    const auto incremental_result = incremental.query(request);
    assert(incremental_result.rows == result.rows);

    const auto username = mapped.resolve_identity_from_name(
        "Username", "Sunil Kumar Peela");
    assert(username.status == point::IdentityResolutionStatus::Unique);
    assert(username.value == "speela");
    const auto employee = mapped.resolve_identity_from_name(
        "Employee ID", "Peela, Sunil Kumar");
    assert(employee.status == point::IdentityResolutionStatus::Unique);
    assert(employee.value == "E1001");
    const auto email = mapped.resolve_identity_from_name(
        "Email", "Sunil Kumar Peela");
    assert(email.status == point::IdentityResolutionStatus::Unique);
    assert(email.value == "sunil@example.com");

    const auto computer_to_username = mapped.universal_lookup(
        {"Username"}, "PC-001");
    assert(computer_to_username.rows.size() == 1);
    assert(computer_to_username.rows[0][0] == "speela");
    const auto email_to_computer = mapped.universal_lookup(
        {"Computer Name"}, "sunil@example.com");
    assert(email_to_computer.rows.size() == 1);
    assert(email_to_computer.rows[0][0] == "PC-001");
    const auto email_to_employee = mapped.universal_lookup(
        {"Employee ID"}, "john@example.com");
    assert(email_to_employee.rows.size() == 1);
    assert(email_to_employee.rows[0][0] == "E1002");
    const auto computer_to_name = mapped.universal_lookup(
        {"Display Name"}, "PC-001");
    assert(computer_to_name.rows.size() == 1);
    assert(computer_to_name.rows[0][0] == "Peela Sunil Kumar");
    const auto one_employee = mapped.universal_lookup(
        {"Employee ID", "Username", "Email"}, "E1001");
    assert(one_employee.rows.size() == 1);
    assert(one_employee.rows[0][0] == "E1001");
    assert(one_employee.rows[0][1] == "speela");
    assert(one_employee.rows[0][2] == "sunil@example.com");
    const auto missing_lookup = mapped.universal_lookup(
        {"Username"}, "COMPUTER-DOES-NOT-EXIST");
    assert(missing_lookup.rows.empty());

    const auto zero_users = root / "ZeroPaddedUsers.csv";
    const auto zero_devices = root / "UnpaddedDevices.csv";
    {
        std::ofstream out(zero_users);
        out << "Employee ID,Username\n00039929,zero.user\n";
    }
    {
        std::ofstream out(zero_devices);
        out << "Associate Number,Computer Name\n39929,PC-ZERO\n";
    }
    point::Engine zero_engine;
    zero_engine.set_field_synonyms({
        {"Employee ID", {"Associate Number"}}
    });
    zero_engine.load_files({zero_users, zero_devices});
    assert(zero_engine.relationships().size() == 1);
    const auto short_eid = zero_engine.universal_lookup(
        {"Username"}, "39929");
    assert(short_eid.rows.size() == 1);
    assert(short_eid.rows[0][0] == "zero.user");
    point::QueryRequest short_eid_request;
    short_eid_request.lookup_field = "Employee ID";
    short_eid_request.lookup_value = "39929";
    short_eid_request.output_fields = {"Username"};
    const auto direct_short_eid = zero_engine.query(short_eid_request);
    assert(direct_short_eid.rows.size() == 1);
    assert(direct_short_eid.rows[0][0] == "zero.user");
    const auto padded_eid = zero_engine.universal_lookup(
        {"Computer Name"}, "00039929");
    assert(padded_eid.rows.size() == 1);
    assert(padded_eid.rows[0][0] == "PC-ZERO");
    const auto preserved_eid = zero_engine.universal_lookup(
        {"Employee ID"}, "zero.user");
    assert(preserved_eid.rows.size() == 1);
    assert(preserved_eid.rows[0][0] == "00039929");

    const auto group_memberships = root / "GroupMemberships.csv";
    const auto group_catalog = root / "GroupCatalog.csv";
    {
        std::ofstream out(group_memberships);
        out << "Employee ID,Group ID,Group Name\n"
               "E000004,G0106,Legal Standard Access\n";
    }
    {
        std::ofstream out(group_catalog);
        out << "Group ID,Group Name,Group Owner\n"
               "G0106,Legal Standard Access,Legal\n";
    }
    point::Engine group_engine;
    group_engine.load_files({group_memberships, group_catalog});
    assert(!group_engine.relationships().empty());
    point::QueryRequest owner_request;
    owner_request.lookup_field = "Employee ID";
    owner_request.lookup_value = "E000004";
    owner_request.output_fields = {
        "Group ID", "Group Name", "Group Owner"};
    const auto owner_result = group_engine.query(owner_request);
    assert(owner_result.rows.size() == 1);
    assert(owner_result.rows[0][0] == "G0106");
    assert(owner_result.rows[0][2] == "Legal");

    const auto unsafe_groups_left = root / "UnsafeGroupsLeft.csv";
    const auto unsafe_groups_right = root / "UnsafeGroupsRight.csv";
    {
        std::ofstream out(unsafe_groups_left);
        out << "Group ID,Left Value\nG1,A\nG1,B\n";
    }
    {
        std::ofstream out(unsafe_groups_right);
        out << "Group ID,Right Value\nG1,C\nG1,D\n";
    }
    point::Engine unsafe_group_engine;
    unsafe_group_engine.load_files({unsafe_groups_left, unsafe_groups_right});
    assert(unsafe_group_engine.relationships().empty());

    const auto request_assignments = root / "RequestAssignments.csv";
    const auto request_catalog = root / "RequestCatalog.csv";
    {
        std::ofstream out(request_assignments);
        out << "Employee ID,Request Code\nE000004,R-42\n";
    }
    {
        std::ofstream out(request_catalog);
        out << "Request Code,Approver\nR-42,Security Manager\n";
    }
    point::Engine generic_catalog_engine;
    generic_catalog_engine.load_files({request_assignments, request_catalog});
    point::QueryRequest approver_request;
    approver_request.lookup_field = "Employee ID";
    approver_request.lookup_value = "E000004";
    approver_request.output_fields = {"Request Code", "Approver"};
    const auto approver_result = generic_catalog_engine.query(approver_request);
    assert(approver_result.rows.size() == 1);
    assert(approver_result.rows[0][1] == "Security Manager");

    const auto duplicates = root / "DuplicateNames.csv";
    {
        std::ofstream out(duplicates);
        out << "Employee ID,Username,Display Name\n"
               "E2001,alex.one,Alex Lee\n"
               "E2002,alex.two,Lee Alex\n";
    }
    point::Engine ambiguous_engine;
    ambiguous_engine.load_files({duplicates});
    const auto ambiguous = ambiguous_engine.resolve_identity_from_name(
        "Username", "Alex Lee");
    assert(ambiguous.status ==
        point::IdentityResolutionStatus::Ambiguous);
    assert(ambiguous.distinct_matches == 2);

    point::Engine contextual_synonym_engine;
    contextual_synonym_engine.set_field_synonyms({
        {"Department", {"Business Unit"}}
    });

    bool conflict_rejected = false;
    try {
        mapped.set_field_synonyms({
            {"Employee ID", {"Worker Number"}},
            {"User ID", {"Worker Number"}}
        });
    } catch (...) {
        conflict_rejected = true;
    }
    assert(conflict_rejected);

    std::filesystem::remove_all(root);
    std::cout << "schema mapping tests passed\n";
}
