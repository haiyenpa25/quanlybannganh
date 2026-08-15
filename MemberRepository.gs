/**
 * MemberRepository.gs - OOP Repository for Members Domain Entity
 */
class MemberRepository extends BaseRepository {
  constructor() {
    super("Members");
  }

  getMembersList() {
    const rawMembers = this.getAllAsObjects();
    const today = new Date();
    const currentMonth = today.getMonth() + 1;

    return rawMembers.map(item => {
      const dobStr = item["Ngày Sinh"] ? String(item["Ngày Sinh"]).trim() : "";
      let age = null;
      let isBirthdayThisMonth = false;

      if (dobStr) {
        const parts = dobStr.split(/[-/]/);
        if (parts.length === 3) {
          const birthMonth = parseInt(parts[1], 10);
          const birthYear = parseInt(parts[0].length === 4 ? parts[0] : parts[2], 10);
          if (!isNaN(birthYear)) age = today.getFullYear() - birthYear;
          if (birthMonth === currentMonth) isBirthdayThisMonth = true;
        }
      }

      return {
        memberId: String(item["Mã Thành Viên"] || ""),
        fullName: String(item["Họ và Tên"] || ""),
        position: String(item["Chức Danh"] || "Ban viên"),
        term: String(item["Nhiệm Kỳ"] || ""),
        dateOfBirth: dobStr,
        age: age,
        isBirthdayThisMonth: isBirthdayThisMonth,
        phone: String(item["Số Điện Thoại"] || ""),
        address: String(item["Địa Chỉ"] || ""),
        groupId: String(item["Mã Tổ"] || ""),
        groupName: String(item["Tên Tổ"] || ""),
        baptizedDate: String(item["Ngày Báp Têm"] || ""),
        spouseName: String(item["Tên Vợ Chồng"] || ""),
        joinDate: String(item["Ngày Sinh Hoạt"] || ""),
        lastAttendanceDate: String(item["Lần Nhóm Cuối"] || ""),
        status: String(item["Trạng Thái"] || "Đang sinh hoạt"),
        notes: String(item["Ghi Chú"] || ""),
        _rowIndex: item._rowIndex
      };
    });
  }

  findMemberById(memberId) {
    const members = this.getMembersList();
    return members.find(m => m.memberId === memberId) || null;
  }
}
